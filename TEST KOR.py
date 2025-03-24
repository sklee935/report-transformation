import re
import pandas as pd

def is_date_format(s):
    """
    간단히 MM/DD/YYYY 형태인지 확인하는 함수
    """
    return bool(re.match(r'^\d{1,2}/\d{1,2}/\d{4}$', s))

def convert_number(s):
    """
    쉼표와 괄호를 제거해 숫자로 변환하기 위한 함수
    예) "5,200.00"   -> 5200.00 (float)
       "(5,837.50)" -> -5837.50 (float)
       "" (빈 문자열) -> 0.0 으로 처리
    """
    s = s.strip()
    # 쉼표 제거
    s = s.replace(',', '')
    # 괄호로 묶여 있으면 음수로 처리
    if s.startswith('(') and s.endswith(')'):
        s = '-' + s[1:-1]
    # 빈 문자열이면 0으로 처리 (원하는 방식에 맞게 조정 가능)
    if s == '':
        return 0.0
    return float(s)

# ------------------------------------------------------------------------
# 1) 입력/출력 경로 설정
# ------------------------------------------------------------------------
input_path = r'C:\Users\slee\OneDrive - SBP\Desktop\sungkeun\report tranformation\INPUT BNA FA Disposal Reeb.xls'
output_path = r'C:\Users\slee\OneDrive - SBP\Desktop\sungkeun\report tranformation\Output BNA_FA_Disposal_Reeb.xlsx'

# ------------------------------------------------------------------------
# 2) Excel 파일 읽기
# ------------------------------------------------------------------------
df_raw = pd.read_excel(input_path, header=None)

records = []
current_gl_full = None  # "80-000-10002030" 같이 전체 GL 번호
current_gl_short = None # 최종 테이블에 쓸 번호 (예: "10002030")
asset_desc = None       # 자산 설명 (다음 행의 자산 레코드와 매칭)

# ------------------------------------------------------------------------
# 3) 각 행을 순회하며 필요한 데이터 추출
# ------------------------------------------------------------------------
for idx, row in df_raw.iterrows():
    # (A) 행 전체를 하나의 문자열로 합침 (빈 셀 제외)
    row_str_list = []
    for cell in row:
        if pd.notnull(cell):
            row_str_list.append(str(cell))
    line = " ".join(row_str_list).strip()

    # (B) "Asset GL Acct #:" 행을 만나면 GL 계정번호 추출
    if line.startswith("Asset GL Acct #:"):
        m = re.search(r'Asset GL Acct #:\s*([\d\-]+)', line)
        if m:
            current_gl_full = m.group(1)  # 예: "80-000-10002030"
            # "80-000-10002030" 에서 뒤쪽 "10002030"만 추출
            if len(current_gl_full) >= 7:
                current_gl_short = current_gl_full[7:]
            else:
                current_gl_short = current_gl_full
        continue

    # (C) 불필요한 행(빈 행, Subtotal, Page, Printed 등) 건너뛰기
    if (not line) or line.startswith("Subtotal:") or line.startswith("Page:") or line.startswith("Printed:"):
        continue

    # (D) 자산 설명 행인지 확인
    #     예: "2011 TOYOTA SIENNA" 처럼 날짜/숫자 없는 순수 텍스트
    if re.match(r'^[A-Za-z0-9\s]+$', line) and not re.search(r'\d{1,2}/\d{1,2}/\d{4}', line):
        asset_desc = line
        continue

    # (E) 자산 정보 행인지 확인
    #     예: "1 10/15/2021 03/05/2024 11,000.00 11,000.00 100.00 100.00"
    #     또는 "20-000480 10/15/2021 02/02/2024 5200.00 5200.00 150.00 150.00"
    fields = re.split(r'\s+', line)
    if len(fields) >= 7:
        # 2번째와 3번째 필드가 날짜(MM/DD/YYYY)인지 검사
        if is_date_format(fields[1]) and is_date_format(fields[2]):
            asset_id      = fields[0]
            placed_date   = fields[1]
            disposal_date = fields[2]
            cost_plus     = convert_number(fields[3])
            ltd_depr      = convert_number(fields[4])
            net_proceeds  = convert_number(fields[5])
            realized_gain = convert_number(fields[6])

            record = {
                'Asset ID': asset_id,
                'Asset Description': asset_desc if asset_desc else "",
                'Placed In Service': placed_date,
                'Disposal Date': disposal_date,
                'Cost Plus Exp. of Sale': cost_plus,    # float
                'LTD Depr & S179/A & AFYD': ltd_depr,   # float
                'Net Proceeds': net_proceeds,           # float
                'Realized Gain (Loss)': realized_gain,  # float
                'Asset GL Acct #': current_gl_short
            }
            records.append(record)
            asset_desc = None  # 자산 설명 초기화

# ------------------------------------------------------------------------
# 4) 결과를 DataFrame으로 만들고 Excel로 저장
# ------------------------------------------------------------------------
df_result = pd.DataFrame(records)
df_result.to_excel(output_path, index=False)

print("결과 파일이 저장되었습니다:", output_path)
