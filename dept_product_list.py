import re
import csv
import os
import sys
from datetime import datetime
from tkinter import filedialog, messagebox

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


def parse_items(item_str: str) -> list:
    """'商品名 x数量 | 商品名 x数量' を [(商品名, 数量), ...] に変換"""
    result = []
    for part in item_str.split(' | '):
        part = part.strip()
        m = re.match(r'^(.+)\s+x(\d+)$', part)
        if m:
            result.append((m.group(1).strip(), int(m.group(2))))
    return result


def run() -> None:
    csv_path = filedialog.askopenfilename(
        title='部署別CSVを選択してください',
        filetypes=[('CSVファイル', '*.csv'), ('すべてのファイル', '*.*')],
    )
    if not csv_path:
        return

    dept_data = {}  # {部署名: {商品名: 数量}}

    try:
        with open(csv_path, encoding='utf-8-sig', newline='') as f:
            reader = csv.DictReader(f)
            if reader.fieldnames:
                reader.fieldnames = [k.lstrip('﻿').strip() for k in reader.fieldnames]
            cols = set(reader.fieldnames or [])
            if '部署名' not in cols or '商品内容' not in cols:
                messagebox.showerror(
                    'エラー',
                    '部署別CSVを選択してください。\n'
                    'WordPressの「部署別CSVをダウンロード」で出力したファイルを選択してください。'
                )
                return
            for row in reader:
                dept = row.get('部署名', '').strip()
                item_str = row.get('商品内容', '').strip()
                if not dept or not item_str:
                    continue
                if dept not in dept_data:
                    dept_data[dept] = {}
                for name, qty in parse_items(item_str):
                    dept_data[dept][name] = dept_data[dept].get(name, 0) + qty
    except Exception as e:
        messagebox.showerror('エラー', f'CSV読み込みエラー:\n{e}')
        return

    if not dept_data:
        messagebox.showerror('エラー', 'CSVにデータがありません')
        return

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    header_font = Font(bold=True, color='FFFFFF')
    header_fill = PatternFill(fill_type='solid', fgColor='2563EB')
    thin = Side(style='thin')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for dept in sorted(dept_data.keys()):
        sheet_name = re.sub(r'[\\/*?:\[\]]', '', dept)[:31]
        ws = wb.create_sheet(title=sheet_name)

        ws.merge_cells('A1:B1')
        ws['A1'].value = dept
        ws['A1'].font = Font(bold=True, size=12)
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[1].height = 50

        ws['A2'].value = '商品名'
        ws['B2'].value = '数量'
        for cell in (ws['A2'], ws['B2']):
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = border
        ws.row_dimensions[2].height = 25

        for i, (name, qty) in enumerate(sorted(dept_data[dept].items()), start=3):
            c1 = ws.cell(row=i, column=1, value=name)
            c1.border = border
            c2 = ws.cell(row=i, column=2, value=qty)
            c2.alignment = Alignment(horizontal='center', vertical='center')
            c2.border = border
            ws.row_dimensions[i].height = 20

        ws.column_dimensions['A'].width = 40
        ws.column_dimensions['B'].width = 8

    now = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_path = os.path.join(os.path.dirname(csv_path), f'部署別商品一覧_{now}.xlsx')

    try:
        wb.save(output_path)
    except Exception as e:
        messagebox.showerror('エラー', f'Excel出力エラー:\n{e}')
        return

    messagebox.showinfo('完了', f'出力完了:\n{output_path}')
    os.startfile(output_path)


if __name__ == '__main__':
    import tkinter as tk
    root = tk.Tk()
    root.withdraw()
    run()
    root.destroy()
