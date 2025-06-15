import tkinter as tk
from tkinter import messagebox
import datetime
import requests
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
import io
import os

def run_extract_allocation(company_name, api_key, bgn_de, end_de):
    def get_corp_code_by_name(company_name, api_key):
        url = f'https://opendart.fss.or.kr/api/corpCode.xml?crtfc_key={api_key}'
        response = requests.get(url)
        if response.status_code != 200:
            messagebox.showerror("오류", '고유번호 API 요청 실패')
            return None
        with zipfile.ZipFile(io.BytesIO(response.content)) as z:
            xml_filename = z.namelist()[0]
            with z.open(xml_filename) as xml_file:
                xml_data = xml_file.read()
        root = ET.fromstring(xml_data)
        for corp in root.findall('list'):
            corp_name = corp.find('corp_name').text
            if corp_name == company_name:
                return corp.find('corp_code').text
        messagebox.showerror("오류", f'"{company_name}"에 해당하는 고유번호를 찾을 수 없습니다.')
        return None

    def get_report_no(corp_code, api_key):
        url = (f'https://opendart.fss.or.kr/api/list.json?crtfc_key={api_key}'
               f'&corp_code={corp_code}&bgn_de={bgn_de}&end_de={end_de}'
               f'&page_no=1&page_count=100')
        response = requests.get(url)
        if response.status_code != 200:
            messagebox.showerror("오류", '공시리스트 API 요청 실패')
            return None
        data = response.json()
        if 'list' not in data:
            messagebox.showerror("오류", '공시 리스트가 없습니다.')
            return None
        for item in data['list']:
            if item['report_nm'].strip() == "증권발행실적보고서":
                return item['rcept_no']
        messagebox.showerror("오류", '"증권발행실적보고서"가 없습니다.')
        return None

    def download_document_by_rcept_no(rcept_no, api_key):
        url = f'https://opendart.fss.or.kr/api/document.xml?crtfc_key={api_key}&rcept_no={rcept_no}'
        response = requests.get(url)
        if response.status_code != 200:
            messagebox.showerror("오류", '보고서 파일 다운로드 실패')
            return None
        zip_filename = f'document_{rcept_no}.zip'
        with open(zip_filename, 'wb') as f:
            f.write(response.content)
        return zip_filename

    def extract_allocation_table(zip_filename, company_name, rcept_no):
        with zipfile.ZipFile(zip_filename, 'r') as zip_ref:
            xml_filename = zip_ref.namelist()[0]
            zip_ref.extract(xml_filename)
        with open(xml_filename, 'r', encoding='utf-8') as f:
            xml_data = f.read()
        root = ET.fromstring(xml_data)
        section_3 = None
        for section in root.findall('.//SECTION-3[@ACLASS="MANDATORY"]'):
            title = section.find('TITLE')
            if title is not None and title.text and '청약 및 배정현황' in title.text:
                section_3 = section
                break
        if section_3 is None:
            messagebox.showerror("오류", '해당 섹션을 찾을 수 없습니다.')
            os.remove(xml_filename)
            return False
        table_group = section_3.find('TABLE-GROUP')
        if table_group is None:
            messagebox.showerror("오류", 'TABLE-GROUP 태그를 찾을 수 없습니다.')
            os.remove(xml_filename)
            return False
        tables = table_group.findall('TABLE')
        if len(tables) < 2:
            messagebox.showerror("오류", '적절한 TABLE이 없습니다.')
            os.remove(xml_filename)
            return False
        target_table = tables[1]
        thead = target_table.find('THEAD')
        header_rows = thead.findall('TR')
        headers_by_row = []
        for tr in header_rows:
            row_headers = []
            for th in tr.findall('TH'):
                text = th.text.strip() if th.text else ''
                colspan = int(th.get('COLSPAN', '1'))
                row_headers.extend([text] * colspan)
            headers_by_row.append(row_headers)
        def merge_headers(header_rows):
            max_len = max(len(r) for r in header_rows)
            for r in header_rows:
                if len(r) < max_len:
                    r.extend([''] * (max_len - len(r)))
            merged = []
            for col in range(max_len):
                parts = [header_rows[row][col] for row in range(len(header_rows)) if header_rows[row][col]]
                merged.append(' - '.join(parts))
            return merged
        final_headers = merge_headers(headers_by_row)
        tbody = target_table.find('TBODY')
        data_rows = []
        for tr in tbody.findall('TR'):
            row = []
            for cell in tr:
                if cell.tag in ('TD', 'TE', 'TH'):
                    txt = cell.text.strip() if cell.text else ''
                    row.append(txt)
            if len(row) < len(final_headers):
                row.extend([''] * (len(final_headers) - len(row)))
            elif len(row) > len(final_headers):
                row = row[:len(final_headers)]
            data_rows.append(row)
        excel_filename = f'{company_name}_증권발행실적_청약및배정현황_{rcept_no}.xlsx'
        df = pd.DataFrame(data_rows, columns=final_headers)
        df.to_excel(excel_filename, index=False)
        messagebox.showinfo("완료", f'엑셀 파일 저장 완료: {excel_filename}')
        os.remove(xml_filename)
        os.remove(zip_filename)
        return True

    # 실행 순서
    corp_code = get_corp_code_by_name(company_name, api_key)
    if corp_code is None:
        return
    report_no = get_report_no(corp_code, api_key)
    if report_no is None:
        return
    zip_filename = download_document_by_rcept_no(report_no, api_key)
    if zip_filename is None:
        return
    extract_allocation_table(zip_filename, company_name, report_no)

# Tkinter GUI 생성
root = tk.Tk()
root.title("DART 증권발행실적보고서 추출")

# 오늘 날짜 및 30일 전 날짜 계산
now = datetime.datetime.now()
default_end = now.strftime('%Y%m%d')
default_bgn = (now - datetime.timedelta(days=30)).strftime('%Y%m%d')

# 라벨 및 입력창
tk.Label(root, text="API KEY:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
entry_apikey = tk.Entry(root, width=40)
entry_apikey.grid(row=0, column=1, padx=5, pady=5, sticky='w')

tk.Label(root, text="종목명:").grid(row=1, column=0, padx=5, pady=5, sticky='w')
entry_company = tk.Entry(root, width=30)
entry_company.grid(row=1, column=1, padx=5, pady=5, sticky='w')

tk.Label(root, text="시작일(YYYYMMDD):").grid(row=2, column=0, padx=5, pady=5, sticky='w')
entry_bgn = tk.Entry(root, width=15)
entry_bgn.grid(row=2, column=1, padx=5, pady=5, sticky='w')
entry_bgn.insert(0, default_bgn)  # 시작일 기본값

tk.Label(root, text="종료일(YYYYMMDD):").grid(row=3, column=0, padx=5, pady=5, sticky='w')
entry_end = tk.Entry(root, width=15)
entry_end.grid(row=3, column=1, padx=5, pady=5, sticky='w')
entry_end.insert(0, default_end)  # 종료일 기본값

def on_submit():
    company = entry_company.get()
    api_key = entry_apikey.get()
    bgn_de = entry_bgn.get()
    end_de = entry_end.get()
    # 날짜 형식 검증
    try:
        datetime.datetime.strptime(bgn_de, '%Y%m%d')
        datetime.datetime.strptime(end_de, '%Y%m%d')
    except ValueError:
        messagebox.showerror("입력 오류", "날짜는 YYYYMMDD 형식으로 입력하세요.")
        return
    if not (company and api_key):
        messagebox.showerror("입력 오류", "종목명과 API KEY를 모두 입력하세요.")
        return
    run_extract_allocation(company, api_key, bgn_de, end_de)

btn = tk.Button(root, text="실행", command=on_submit)
btn.grid(row=4, column=0, columnspan=2, pady=10)

root.mainloop()
