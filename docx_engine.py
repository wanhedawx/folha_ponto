"""
Motor de geração de Folhas de Ponto a partir do template Word (mala direta)
+ período de trabalho escolhido pelo usuário (data inicial/final).
"""
import re
import random
import zipfile
import io
from datetime import date, timedelta

WEEKDAY_PT = ['SEG', 'TER', 'QUA', 'QUI', 'SEX', 'SAB', 'DOM']  # Monday=0
MONTH_PT = ['JAN', 'FEV', 'MAR', 'ABR', 'MAI', 'JUN', 'JUL', 'AGO', 'SET', 'OUT', 'NOV', 'DEZ']

MERGE_FIELDS = [
    'NOME_EMPRESA',
    'DESCRIÇÃO_DO_LOCAL',
    'ENDEREÇO',
    'NUMERO',
    'CIDADE',
    'BAIRRO',
    'MATRICULA',
    'NOME',
    'ADMISSAO',
    'FUNÇÃO',
    'HORARIO',
    'PERIODO',
    'ANO'
]

def _xml_escape(text: str) -> str:
    return (
        str(text)
        .replace('&', '&amp;')
        .replace('<', '&lt;')
        .replace('>', '&gt;')
    )


def _rand_id() -> str:
    return ''.join(random.choice('0123456789ABCDEF') for _ in range(8))


def load_template_xml(template_bytes: bytes) -> tuple[str, dict]:
    """Lê o document.xml do template e simplifica (remove o Fallback VML
    duplicado, mantendo só o desenho moderno). Retorna (xml, outros_arquivos)."""
    z = zipfile.ZipFile(io.BytesIO(template_bytes))
    others = {name: z.read(name) for name in z.namelist() if name != 'word/document.xml'}
    content = z.read('word/document.xml').decode('utf-8')

    alt_start = content.find('<mc:AlternateContent')
    if alt_start != -1:
        alt_end = content.find('</mc:AlternateContent>') + len('</mc:AlternateContent>')
        choice_start = content.find('<mc:Choice', alt_start)
        drawing_start = content.find('<w:drawing>', choice_start)
        drawing_end = content.find('</w:drawing>', drawing_start) + len('</w:drawing>')
        content = content[:alt_start] + content[drawing_start:drawing_end] + content[alt_end:]

    return content, others


def _extract_day_rows_block(content: str):
    """Localiza o bloco contíguo de linhas de dia (ex: '16' 'QUI', '17' 'SEX'...)
    e retorna (start_idx, end_idx, linha_modelo_xml)."""
    weekdays = set(WEEKDAY_PT)
    for m in re.finditer(r'<w:tr .*?</w:tr>', content):
        row = m.group(0)
        texts = re.findall(r'<w:t(?:\s[^>]*)?>(.*?)</w:t>', row)
        if len(texts) >= 2 and texts[0].strip().isdigit() and texts[1].strip() in weekdays:
            first_match = m
            break
    else:
        raise ValueError('Não encontrei as linhas de calendário no template.')

    # modelo de linha "limpo" (sem vMerge) — pega a segunda linha de dia encontrada
    # para evitar pegar a primeira, que pode ter bordas especiais.
    day_rows = []
    for m in re.finditer(r'<w:tr .*?</w:tr>', content):
        row = m.group(0)
        texts = re.findall(r'<w:t(?:\s[^>]*)?>(.*?)</w:t>', row)
        if len(texts) >= 2 and texts[0].strip().isdigit() and texts[1].strip() in weekdays:
            day_rows.append(m)

    start_idx = day_rows[0].start()
    end_idx = day_rows[-1].end()
    template_row = day_rows[1].group(0) if len(day_rows) > 1 else day_rows[0].group(0)
    return start_idx, end_idx, template_row


def _regen_ids(row_xml: str) -> str:
    """Troca todos os w14:paraId / w14:textId por valores novos (evita IDs duplicados)."""
    row_xml = re.sub(r'w14:paraId="[0-9A-Fa-f]{8}"', lambda _: f'w14:paraId="{_rand_id()}"', row_xml)
    row_xml = re.sub(r'w14:textId="[0-9A-Fa-f]{8}"', lambda _: f'w14:textId="{_rand_id()}"', row_xml)
    return row_xml


def _build_day_row(template_row: str, day_num: int, weekday_abbr: str) -> str:
    row = _regen_ids(template_row)
    # troca o primeiro <w:t>...</w:t> (número do dia) e o segundo (dia da semana)
    matches = list(re.finditer(r'(<w:t(?:\s[^>]*)?>)(.*?)(</w:t>)', row))
    if len(matches) < 2:
        raise ValueError('Linha de calendário com estrutura inesperada.')
    day_m, wd_m = matches[0], matches[1]
    row = (
        row[:day_m.start(2)] + str(day_num) + row[day_m.end(2):]
    )
    # recalcula posição do segundo match após a primeira substituição
    matches2 = list(re.finditer(r'(<w:t(?:\s[^>]*)?>)(.*?)(</w:t>)', row))
    wd_m2 = matches2[1]
    row = row[:wd_m2.start(2)] + weekday_abbr + row[wd_m2.end(2):]
    return row


def _replace_period_row(content: str, start_date: date, end_date: date) -> str:
    idx = content.find('w14:paraId="7B22BC94"')
    if idx == -1:
        return content  # já não existe (ex.: chamado 2x) — ignora
    tr_start = content.rfind('<w:tr ', 0, idx)
    tr_end = content.find('</w:tr>', idx) + len('</w:tr>')
    row = content[tr_start:tr_end]

    # formata "16/04/2026 A 15/05/2026"
    new_text = f"{start_date.strftime('%d/%m/%Y')} A {end_date.strftime('%d/%m/%Y')}"

    p_match = re.search(r'(<w:p [^>]*>.*?</w:pPr>)(.*?)(</w:p>)', row, re.S)
    if not p_match:
        return content
    first_run_rpr = re.search(r'<w:r[^>]*><w:rPr>(.*?)</w:rPr>', p_match.group(2), re.S)
    rpr_xml = f'<w:rPr>{first_run_rpr.group(1)}</w:rPr>' if first_run_rpr else ''
    new_run = f'<w:r>{rpr_xml}<w:t>{_xml_escape(new_text)}</w:t></w:r>'
    new_p = p_match.group(1) + new_run + p_match.group(3)
    new_row = row[:p_match.start()] + new_p + row[p_match.end():]

    return content[:tr_start] + new_row + content[tr_end:]


def fill_merge_fields(content: str, values: dict):

    # Substitui os marcadores @@CAMPO@@
    for campo, valor in values.items():
        content = content.replace(
            f"@@{campo}@@",
            _xml_escape(str(valor))
        )

    # Continua substituindo os MergeFields normais («CAMPO»)
    for field in MERGE_FIELDS:
        value = _xml_escape(values.get(field, ""))
        placeholder = f"«{field}»"

        pattern = re.compile(
            r'(<w:t[^>]*>)' + re.escape(placeholder) + r'(</w:t>)'
        )

        content, _ = pattern.subn(
            lambda m: m.group(1) + value + m.group(2),
            content
        )

    return content

def set_period(content: str, start_date: date, end_date: date) -> str:
    # cabeçalho PERÍODO / ANO
    if start_date.month == end_date.month:
        periodo_txt = MONTH_PT[start_date.month - 1]
    else:
        periodo_txt = f'{MONTH_PT[start_date.month - 1]}/{MONTH_PT[end_date.month - 1]}'
    content = content.replace('ABR/MAI', periodo_txt, 1)
    content = content.replace('>2026<', f'>{start_date.year}<', 1)

    # linha "PERIODO TRABALHADO: dd/mm/aaaa A dd/mm/aaaa"
    content = _replace_period_row(content, start_date, end_date)

    # tabela de dias
    start_idx, end_idx, template_row = _extract_day_rows_block(content)
    days = []
    d = start_date
    while d <= end_date:
        days.append(d)
        d += timedelta(days=1)
    new_rows_xml = ''.join(
        _build_day_row(template_row, d.day, WEEKDAY_PT[d.weekday()]) for d in days
    )
    content = content[:start_idx] + new_rows_xml + content[end_idx:]
    return content


def build_docx(base_content: str, others: dict, values: dict, start_date: date, end_date: date) -> bytes:
    content = fill_merge_fields(base_content, values)
    content = set_period(content, start_date, end_date)

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, data in others.items():
            z.writestr(name, data)
        z.writestr('word/document.xml', content.encode('utf-8'))
    return buf.getvalue()
