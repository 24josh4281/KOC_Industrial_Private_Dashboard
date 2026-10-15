"""Make a brand-consistent v17 view from the validated v16 snapshot.

Only presentation sheets and chart paint are restyled. The workbook's values,
formulas, source sheets, charts and hyperlinks remain available for QA.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import zipfile
from copy import copy
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


ROOT = Path(__file__).resolve().parent
BRAND = json.loads((ROOT / 'brand.json').read_text(encoding='utf-8'))
COLOR = {key: value.lstrip('#').upper() for key, value in BRAND.items()
         if isinstance(value, str) and value.startswith('#')}
NAVY = COLOR['navy']
BLUE = COLOR['blue']
DARK_BLUE = COLOR['dark_blue']
PANEL = COLOR['panel']
GOVERNING = COLOR['governing_fill']
LINE = COLOR['governing_line']
INK = COLOR['ink']
MUTED = COLOR['muted']
WHITE = 'FFFFFF'
LIGHT_TEXT = 'D2DCFA'


def fill(hex_color):
    return PatternFill('solid', fgColor=hex_color)


def restyle_card(ws, first_col):
    for row in ws.iter_rows(min_row=5, max_row=8,
                            min_col=first_col, max_col=first_col + 2):
        for cell in row:
            cell.fill = fill(NAVY)
            if cell.row == 5:
                cell.border = Border(top=Side(style='medium', color=BLUE))
    for row, size, color, bold in ((5, 10, LIGHT_TEXT, True),
                                   (6, 24, WHITE, True),
                                   (8, 9, LIGHT_TEXT, False)):
        cell = ws.cell(row, first_col)
        cell.font = Font(name='Malgun Gothic', size=size, color=color, bold=bold)
        cell.alignment = Alignment(vertical='center', wrap_text=(row == 8))
        if row == 6:
            cell.number_format = '#,##0'


def restyle_chart(chart, color):
    chart.style = 10
    chart.legend = None
    chart.y_axis.scaling.min = 0
    chart.y_axis.numFmt = '#,##0'
    # Long vertical axis titles are clipped in Excel and in exported previews.
    # The adjacent sheet headings/notes already state that counts are source rows.
    chart.y_axis.title = None
    chart.x_axis.title = None
    for series in chart.series:
        series.graphicalProperties.solidFill = color
        series.graphicalProperties.line.noFill = True


def restyle_dashboard(ws, chart_colors):
    ws.sheet_properties.tabColor = NAVY
    ws.sheet_view.showGridLines = False
    for row in ws.iter_rows(min_row=1, max_row=2, min_col=1, max_col=16):
        for cell in row:
            cell.fill = fill(NAVY)
            if cell.row == 1:
                cell.border = Border(top=Side(style='medium', color=BLUE))
    ws['A1'].font = Font(name='Malgun Gothic', size=20, bold=True, color=WHITE)
    ws['A1'].alignment = Alignment(vertical='center')
    ws['A3'].font = Font(name='Malgun Gothic', size=10, color=MUTED)
    for first_col in (1, 5, 9, 13):
        restyle_card(ws, first_col)
    for address in ('A10', 'A29'):
        if ws[address].value:
            ws[address].fill = fill(PANEL)
            ws[address].font = Font(name='Malgun Gothic', size=12,
                                    bold=True, color=NAVY)
            ws[address].border = Border(left=Side(style='medium', color=BLUE))
    for address in ('I11', 'I30', 'A34'):
        if ws[address].value and not str(ws[address].value).startswith('='):
            ws[address].fill = fill(GOVERNING)
            ws[address].font = Font(name='Malgun Gothic', size=10, color=INK)
            ws[address].border = Border(left=Side(style='medium', color=BLUE),
                                        top=Side(style='thin', color=LINE),
                                        bottom=Side(style='thin', color=LINE),
                                        right=Side(style='thin', color=LINE))
            ws[address].alignment = Alignment(wrap_text=True, vertical='center')
    for chart, color in zip(ws._charts, chart_colors, strict=True):
        restyle_chart(chart, color)


def restyle_workbook(source: Path, output: Path):
    wb = load_workbook(source)
    if len(wb.sheetnames) != 77:
        raise ValueError(f'예상한 v16 시트 77개가 아닙니다: {len(wb.sheetnames)}')
    targets = {
        '통계 대시보드': (BLUE, BLUE, DARK_BLUE),
        '진행현황 대시보드': (DARK_BLUE, BLUE),
        '업무 운영 대시보드': (DARK_BLUE, BLUE, BLUE),
    }
    for name, colors in targets.items():
        restyle_dashboard(wb[name], colors)
    # The long ORS status names and reviewer names need horizontal label space.
    wb['진행현황 대시보드']._charts[0].type = 'bar'
    wb['업무 운영 대시보드']._charts[2].type = 'bar'
    operations = wb['업무 운영 대시보드']
    for cell, short_label in {'A11': '조치 입력', 'D11': '의견 입력',
                              'G11': '안내 입력', 'J11': '보완 단계',
                              'A27': '우선순위', 'A39': '검토기관'}.items():
        operations[cell] = short_label
        operations[cell].fill = fill(PANEL)
        operations[cell].font = Font(name='Malgun Gothic', size=10,
                                     bold=True, color=NAVY)
        operations[cell].border = Border(left=Side(style='medium', color=BLUE))
        operations[cell].alignment = Alignment(vertical='center')

    queue = wb['담당자 작업현황']
    queue.sheet_properties.tabColor = BLUE
    queue.sheet_view.showGridLines = False
    for cell in queue[1]:
        cell.fill = fill(NAVY)
        cell.font = Font(name='Malgun Gothic', size=10, bold=True, color=WHITE)
        cell.alignment = Alignment(horizontal='center', vertical='center',
                                   wrap_text=True)
    for row in queue.iter_rows(min_row=2):
        for cell in row:
            if cell.hyperlink:
                font = copy(cell.font)
                font.color = DARK_BLUE
                cell.font = font

    requirements = wb['요구사항 점검']
    requirements.sheet_properties.tabColor = MUTED
    for cell in requirements[1]:
        cell.fill = fill(NAVY)
        cell.font = Font(name='Malgun Gothic', size=10, bold=True, color=WHITE)
        cell.alignment = Alignment(horizontal='center', vertical='center',
                                   wrap_text=True)

    output.parent.mkdir(parents=True, exist_ok=True)
    temporary = output.with_suffix('.building.xlsx')
    try:
        wb.save(temporary)
        with zipfile.ZipFile(temporary) as archive:
            bad = archive.testzip()
            charts = sum(name.startswith('xl/charts/chart') and name.endswith('.xml')
                         for name in archive.namelist())
        if bad or charts != 12:
            raise ValueError(f'Excel 차트/압축 검사 실패: {bad}, 차트 {charts}')
        os.replace(temporary, output)
    finally:
        wb.close()
    return {'sheets': 77, 'charts': charts}


def build(v16: Path, output: Path):
    output.mkdir(parents=True, exist_ok=True)
    for sub in ('data', 'details'):
        shutil.copytree(v16 / sub, output / sub, dirs_exist_ok=True)
    for name in ('statistics.html', 'progress.html', 'case.html', 'work.html',
                 'app.js', 'app.css'):
        content = (v16 / name).read_text(encoding='utf-8')
        content = content.replace('외부사업_통합관리_v15.xlsx', '외부사업_통합관리_v17.xlsx')
        content = content.replace('외부사업_통합관리_v16.xlsx', '외부사업_통합관리_v17.xlsx')
        content = content.replace('외부사업_진행현황_v16.csv', '외부사업_진행현황_v17.csv')
        content = content.replace('외부사업_담당자_작업현황_v16.csv',
                                  '외부사업_담당자_작업현황_v17.csv')
        if name.endswith('.html'):
            content = content.replace('<link rel="stylesheet" href="app.css">',
                                      '<link rel="stylesheet" href="app.css"><link rel="stylesheet" href="brand.css">')
        (output / name).write_text(content, encoding='utf-8')
    if (ROOT / 'brand.css').resolve() != (output / 'brand.css').resolve():
        shutil.copy2(ROOT / 'brand.css', output / 'brand.css')
    excel = restyle_workbook(v16 / '외부사업_통합관리_v16.xlsx',
                             output / '외부사업_통합관리_v17.xlsx')
    prior = json.loads((v16 / 'manifest.json').read_text(encoding='utf-8'))
    digest = hashlib.sha256(prior['version'].encode('ascii'))
    for name in ('build_v17.py', 'brand.css', 'brand.json'):
        digest.update((ROOT / name).read_bytes())
    manifest = {**prior, 'version': digest.hexdigest()[:16],
                'v16_version': prior['version'], 'brand_profile': BRAND['profile'],
                'workbook_sheets': excel['sheets']}
    (output / 'manifest.json').write_text(
        json.dumps(manifest, ensure_ascii=False, separators=(',', ':')),
        encoding='utf-8')
    return {'output': str(output), 'excel': excel, 'manifest': manifest}


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--v16', required=True, type=Path)
    parser.add_argument('--output', required=True, type=Path)
    args = parser.parse_args()
    print(json.dumps(build(args.v16, args.output), ensure_ascii=False))
