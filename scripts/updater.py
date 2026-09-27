import os
import sys
import time
import datetime
import requests
import pandas as pd
import openpyxl
from io import StringIO

# 偽裝成標準瀏覽器，避免被證交所/期交所阻擋
HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept-Language': 'zh-TW,zh;q=0.9,en-US;q=0.8,en;q=0.7'
}

def get_current_date():
    """取得台灣時區 (UTC+8) 的日期字串"""
    # GitHub 伺服器為 UTC，需手動增加 8 小時
    tw_time = datetime.datetime.utcnow() + datetime.timedelta(hours=8)
    return tw_time.strftime("%Y%m%d"), tw_time.strftime("%Y/%m/%d")

def fetch_twse_data(date_str):
    """
    抓取 TWSE：大盤成交金額、三大法人買賣超、融資餘額
    """
    print(f"[TWSE] 開始抓取 {date_str} 盤後資料...")
    
    # 1. 大盤成交量 (MI_INDEX)
    url_index = f"https://www.twse.com.tw/rwd/zh/afterTrading/MI_INDEX?date={date_str}&type=MS&response=json"
    r_index = requests.get(url_index, headers=HEADERS, timeout=15).json()
    if r_index.get('stat') != 'OK':
        print(f"[TWSE] {date_str} 非交易日或證交所尚未公布資料。")
        return None

    vol = None
    for row in r_index.get('tables', [{}])[0].get('data', []):
        if '總計' in row[0]:
            vol = float(row[1].replace(',', ''))
            break
            
    time.sleep(3)

    # 2. 三大法人買賣超金額 (BFIAMU)
    url_fund = f"https://www.twse.com.tw/rwd/zh/fund/BFIAMU?date={date_str}&response=json"
    r_fund = requests.get(url_fund, headers=HEADERS, timeout=15).json()
    foreign_buy, trust_buy, dealer_buy = 0.0, 0.0, 0.0
    if r_fund.get('stat') == 'OK':
        for row in r_fund.get('data', []):
            diff = float(row[3].replace(',', ''))
            if '外資及陸資' in row[0]:
                foreign_buy = diff
            elif '投信' in row[0]:
                trust_buy = diff
            elif '自營商' in row[0]:
                dealer_buy += diff

    time.sleep(3)

    # 3. 融資餘額 (MI_MARGN)
    url_margin = f"https://www.twse.com.tw/rwd/zh/marginTrading/MI_MARGN?date={date_str}&selectType=MS&response=json"
    r_margin = requests.get(url_margin, headers=HEADERS, timeout=15).json()
    margin_balance = None
    if r_margin.get('stat') == 'OK':
        for row in r_margin.get('tables', [{}])[0].get('data', []):
            if '融資金額(仟元)' in row[0]:
                # 換算為元 (仟元 * 1000)
                margin_balance = float(row[5].replace(',', '')) * 1000
                break

    return {
        '成交量': vol,
        '融資餘額': margin_balance,
        '外資': foreign_buy,
        '投信': trust_buy,
        '自營': dealer_buy
    }

def fetch_taifex_data(date_str_formatted):
    """
    抓取 TAIFEX：外資期貨大小、前五/十大留倉、選PCR、外選淨金額、韭菜指數、未平倉口數
    """
    print(f"[TAIFEX] 開始抓取 {date_str_formatted} 期權資料...")
    query_date = date_str_formatted.replace('-', '/')
    
    # 預設回傳字典
    data = {
        '外資期貨大小': None,
        '前五大交易人留倉': None,
        '前十大交易人留倉': None,
        '選PCR': None,
        '外選': None,
        '韭菜指數': None,
        '未平倉口數': None
    }

    # 1. 選擇權 PCR (Put/Call Ratio)
    try:
        url_pcr = f"https://www.taifex.com.tw/cht/3/pcRatio?queryStartDate={query_date}&queryEndDate={query_date}"
        r_pcr = requests.get(url_pcr, headers=HEADERS, timeout=15)
        tables = pd.read_html(StringIO(r_pcr.text))
        if tables and len(tables[0]) > 0:
            data['選PCR'] = float(tables[0].iloc[0]['買賣權未平倉量比率%'])
    except Exception as e:
        print(f"[Warning] PCR 抓取失敗: {e}")

    time.sleep(2)

    # 2. 期貨三大法人未平倉 (台指期外資淨未平倉)
    try:
        url_fut = "https://www.taifex.com.tw/cht/3/futContractsDate"
        params_tx = {'queryDate': query_date, 'contractId': 'TX'}
        r_fut = requests.post(url_fut, data=params_tx, headers=HEADERS, timeout=15)
        fut_tables = pd.read_html(StringIO(r_fut.text))
        for t in fut_tables:
            # 尋找包含多空淨額欄位的表格
            cols = [str(c) for c in t.columns]
            if any('身份別' in c for c in cols):
                for _, row in t.iterrows():
                    row_str = ' '.join(row.astype(str))
                    if '外資' in row_str:
                        # 擷取多空淨額口數
                        val = str(row.iloc[-2]).replace(',', '').strip()
                        data['外資期貨大小'] = int(val)
                        break
    except Exception as e:
        print(f"[Warning] 期貨三大法人抓取失敗: {e}")

    time.sleep(2)

    # 3. 臺指期全市場未平倉口數 (TX 未平倉)
    try:
        url_tx_oi = "https://www.taifex.com.tw/cht/3/futDailyMarketReport"
        params_oi = {'queryDate': query_date, 'MarketCode': '0', 'commodity_id': 'TX'}
        r_oi = requests.post(url_tx_oi, data=params_oi, headers=HEADERS, timeout=15)
        oi_tables = pd.read_html(StringIO(r_oi.text))
        for t in oi_tables:
            if '未沖銷契約量' in [str(c) for c in t.columns]:
                row_subtotal = t[t.iloc[:, 0].astype(str).str.contains('小計', na=False)]
                if not row_subtotal.empty:
                    data['未平倉口數'] = int(str(row_subtotal.iloc[0]['未沖銷契約量']).replace(',', ''))
                    break
    except Exception as e:
        print(f"[Warning] 期貨未平倉口數抓取失敗: {e}")

    time.sleep(2)

    # 4. 前五大與前十大交易人留倉 (TX 所有契約買賣差額)
    try:
        url_large = "https://www.taifex.com.tw/cht/3/largeTraderFutQry"
        params_lg = {'queryDate': query_date, 'contractId': 'TX'}
        r_lg = requests.post(url_large, data=params_lg, headers=HEADERS, timeout=15)
        lg_tables = pd.read_html(StringIO(r_lg.text))
        for t in lg_tables:
            row_all = t[t.iloc[:, 1].astype(str).str.contains('所有契約', na=False)]
            if not row_all.empty:
                # 買方前五大 - 賣方前五大
                top5_buy = int(str(row_all.iloc[0, 2]).replace(',', ''))
                top5_sell = int(str(row_all.iloc[0, 6]).replace(',', ''))
                data['前五大交易人留倉'] = top5_buy - top5_sell

                # 買方前十大 - 賣方前十大
                top10_buy = int(str(row_all.iloc[0, 4]).replace(',', ''))
                top10_sell = int(str(row_all.iloc[0, 8]).replace(',', ''))
                data['前十大交易人留倉'] = top10_buy - top10_sell
                break
    except Exception as e:
        print(f"[Warning] 大額交易人留倉抓取失敗: {e}")

    time.sleep(2)

    # 5. 外資選擇權契約金額淨額 (買權賣權淨多空合計金額)
    try:
        url_opt = "https://www.taifex.com.tw/cht/3/callsAndPutsDate"
        params_opt = {'queryDate': query_date}
        r_opt = requests.post(url_opt, data=params_opt, headers=HEADERS, timeout=15)
        opt_tables = pd.read_html(StringIO(r_opt.text))
        for t in opt_tables:
            if '外資' in ' '.join(t.iloc[:, 0].astype(str)):
                # 買權多空差額 + 賣權多空差額（單位：千元轉為元）
                call_net = float(str(t.iloc[2, -1]).replace(',', '')) * 1000
                put_net = float(str(t.iloc[5, -1]).replace(',', '')) * 1000
                data['外選'] = call_net + put_net
                break
    except Exception as e:
        print(f"[Warning] 外資選擇權抓取失敗: {e}")

    time.sleep(2)

    # 6. 韭菜指數 (散戶小台多空比)
    try:
        params_mtx = {'queryDate': query_date, 'contractId': 'MTX'}
        r_mtx = requests.post(url_fut, data=params_mtx, headers=HEADERS, timeout=15)
        mtx_tables = pd.read_html(StringIO(r_mtx.text))
        
        # 散戶淨部位 = -1 * (三大法人淨部位合計)
        inst_net_oi = 0
        for t in mtx_tables:
            cols = [str(c) for c in t.columns]
            if any('身份別' in c for c in cols):
                for _, row in t.iterrows():
                    val = str(row.iloc[-2]).replace(',', '').strip()
                    if val.lstrip('-').isdigit():
                        inst_net_oi += int(val)

        # 取得小台全市場未沖銷量
        params_mtx_oi = {'queryDate': query_date, 'MarketCode': '0', 'commodity_id': 'MTX'}
        r_mtx_oi = requests.post(url_tx_oi, data=params_mtx_oi, headers=HEADERS, timeout=15)
        mtx_oi_tables = pd.read_html(StringIO(r_mtx_oi.text))
        mtx_total_oi = 1
        for t in mtx_oi_tables:
            if '未沖銷契約量' in [str(c) for c in t.columns]:
                row_subtotal = t[t.iloc[:, 0].astype(str).str.contains('小計', na=False)]
                if not row_subtotal.empty:
                    mtx_total_oi = int(str(row_subtotal.iloc[0]['未沖銷契約量']).replace(',', ''))
                    break

        retail_net_oi = -inst_net_oi
        jiucai_ratio = (retail_net_oi / mtx_total_oi) * 100
        data['韭菜指數'] = f"{jiucai_ratio:.2f}%"
    except Exception as e:
        print(f"[Warning] 韭菜指數抓取失敗: {e}")

    return data

def update_excel_file(file_path, row_data):
    """
    開啟先行指標.xlsx，比對日期回填各欄位數值
    """
    if not os.path.exists(file_path):
        print(f"[Error] 找不到檔案: {file_path}")
        return False

    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active  # 使用第一個主工作表

    target_date = row_data['日期']
    updated = False

    # 遍歷第一欄比對日期
    for row in range(2, sheet.max_row + 1):
        cell_date = str(sheet.cell(row=row, column=1).value or '').strip()
        if target_date in cell_date:
            # 依主表欄位依序填入
            sheet.cell(row=row, column=2, value=row_data.get('成交量'))
            sheet.cell(row=row, column=3, value=row_data.get('融資餘額'))
            sheet.cell(row=row, column=4, value=row_data.get('外資'))
            sheet.cell(row=row, column=5, value=row_data.get('投信'))
            sheet.cell(row=row, column=6, value=row_data.get('自營'))
            sheet.cell(row=row, column=7, value=row_data.get('外資期貨大小'))
            sheet.cell(row=row, column=8, value=row_data.get('前五大交易人留倉'))
            sheet.cell(row=row, column=9, value=row_data.get('前十大交易人留倉'))
            sheet.cell(row=row, column=10, value=row_data.get('選PCR'))
            sheet.cell(row=row, column=11, value=row_data.get('外選'))
            sheet.cell(row=row, column=12, value=row_data.get('韭菜指數'))
            sheet.cell(row=row, column=13, value=row_data.get('未平倉口數'))
            updated = True
            print(f"[Excel] 成功更新 {target_date} 資料列 (Row {row})。")
            break

    # 若檔案預排日期沒有今天，則直接在最後一行追加新資料
    if not updated:
        new_row = [
            target_date,
            row_data.get('成交量'),
            row_data.get('融資餘額'),
            row_data.get('外資'),
            row_data.get('投信'),
            row_data.get('自營'),
            row_data.get('外資期貨大小'),
            row_data.get('前五大交易人留倉'),
            row_data.get('前十大交易人留倉'),
            row_data.get('選PCR'),
            row_data.get('外選'),
            row_data.get('韭菜指數'),
            row_data.get('未平倉口數')
        ]
        sheet.append(new_row)
        print(f"[Excel] 日期未預排，已於末端追加新列：{target_date}。")

    wb.save(file_path)
    return True

def main():
    date_compact, date_slash = get_current_date()
    print(f"=== 執行每日盤後先行指標更新：{date_slash} ===")

    # 1. 抓取證交所現貨與信用交易資料
    twse_data = fetch_twse_data(date_compact)
    if not twse_data:
        print("[Info] 今日非交易日或無資料更新，程式優雅結束。")
        sys.exit(0)

    # 2. 抓取期交所期權與散戶籌碼資料
    taifex_data = fetch_taifex_data(date_slash)

    # 3. 整合所有指標
    row_data = {
        '日期': date_slash,
        **twse_data,
        **taifex_data
    }

    # 4. 回填至 Excel 檔案
    excel_path = "先行指標.xlsx"
    update_excel_file(excel_path, row_data)
    print("=== 先行指標更新完成 ===")

if __name__ == '__main__':
    main()
