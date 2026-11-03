"""シート名と列番号の定義。列番号は A列=1 で数える。"""

ACCOUNT_2219 = 2219
ACCOUNT_TEMP_TAX = 384  # 仮払消費税

SOURCE_SHEET = "科目２０００～５０００未満"
TARGET_SHEET = "0384仮払い消費税"
HEADER_ROWS = 3  # 見出し2行と前月残の行

# 「科目２０００～５０００未満」シート
SRC_ACCOUNT = 1  # A列: 元帳の科目
SRC_VOUCHER = 4  # D列: 伝票№
SRC_SUBJECT = 5  # E列: 科目
SRC_DEBIT = 8  # H列: 借方
SRC_CREDIT = 9  # I列: 貸方
SRC_PARENT_VOUCHER = 12  # L列: 親伝票№

# 「0384仮払い消費税」シート
TGT_VOUCHER = 4  # D列: 伝票№
TGT_SUBJECT_CODE = 5  # E列: 相手科目
TGT_RATE = 8  # H列: 税率
TGT_DEBIT = 10  # J列: 借方
TGT_DEBIT_SUBJECT = 11  # K列: 借方の科目
TGT_CREDIT = 15  # O列: 貸方
TGT_CREDIT_SUBJECT = 16  # P列: 貸方の科目
TGT_TAX_CLASS = 21  # U列: 課税区分
TGT_COPY_LAST_COLUMN = 7  # 追加行に複写する最後の列（G列）

# 税率ごとの税額列
TGT_DEBIT_TAX_COLUMNS = {5: 12, 8: 13, 10: 14}  # L, M, N列
TGT_CREDIT_TAX_COLUMNS = {5: 17, 8: 18, 10: 19}  # Q, R, S列

# 課税区分から考えられる税率。「外」は金額の比率で判定する
TAX_CLASS_PERCENTS = {
    "課10": [10],
    "免10": [10],
    "課08": [8],
    "課8": [8],
    "免08": [8],
    "課05": [5],
    "外": [10, 8],
}

# 税額のみ入力する行で、H列の税率と照合する課税区分
TAXABLE_CLASS_PERCENTS = {"課10": 10, "課08": 8, "課8": 8}

RED = "FFFF0000"
