import pandas as pd
import numpy as np
import os
import re
import io  # 改用 io.StringIO
import subprocess  # 用於從 GitHub 拉取最新版本

# 0214版，修改 商品金額總 =pd.NamedAgg(column="商品金額", aggfunc=lambda x: x.sum().abs()) 改成先sum 再 abs
# 0216_v1版 Debug 商品金額總=pd.NamedAgg(column="商品金額", aggfunc=lambda x: abs(x.sum()))
# 0216_v2版 Debug 油品類數量總=pd.NamedAgg(column="油品類數量", aggfunc=lambda x: abs(x.sum())),

print("================================")
print(" 中油車隊卡 交易彙整工具")
print("================================")
print()
print("當前版本號 1201_v5")

# 獲取當前目錄
current_dir = os.path.dirname(os.path.abspath(__file__))
input_file = os.path.join(current_dir, "交易查詢會總表.csv")


def update_from_github():
    """從 GitHub 拉取最新版本（取代原本的 .command 啟動檔，Windows / macOS 皆可用）。"""
    print("🔄 從 GitHub 拉取最新版本...")
    try:
        result = subprocess.run(
            ["git", "pull", "--ff-only"],
            cwd=current_dir,
            capture_output=True,
            text=True,
        )
        if result.returncode == 0:
            print("✅ 已更新到最新版本")
        else:
            print("⚠️  git pull 失敗（可能無網路或有衝突），將使用本地版本繼續執行")
    except FileNotFoundError:
        print("⚠️  找不到 git 指令，將使用本地版本繼續執行")
    except Exception:
        print("⚠️  git pull 發生錯誤，將使用本地版本繼續執行")
    print()


update_from_github()

# 確保 CSV 存在
if not os.path.exists(input_file):
    print(f"❌ 檔案未找到: {input_file}")
    input("請將 '交易查詢會總表.csv' 放在此目錄後再執行，按 Enter 退出...")
    exit()


def detect_and_fix_csv(file_path, encoding="big5"):
    # 檢查是否有 `="..."` 問題（代表 Excel 文字格式異常）
    with open(file_path, "r", encoding=encoding) as f:
        lines = f.readlines()

    # 檢測是否有 `="..."` 問題
    needs_fix = any(re.search(r'=".*?"', line) for line in lines)

    if needs_fix:
        print("⚠️ 檢測到 Excel 文字格式 (`='...'`)，正在修復...")
        
        # 找到欄位名稱行
        header_line_idx = next(i for i, line in enumerate(lines) if "," in line)
        header = lines[header_line_idx].strip()  # 取得欄位名稱
        data_lines = lines[header_line_idx + 1:]  # 取得數據部分

        # 修正 `="..."` 問題
        fixed_lines = [re.sub(r'="(.*?)"', r'\1', line) for line in data_lines]

        # 直接讀取修正後的數據，確保 `index_col=False`
        csv_data = io.StringIO("\n".join([header] + fixed_lines))
        df = pd.read_csv(csv_data, encoding=encoding, index_col=False)

        # 不在這裡解析交易時間，讓主程序統一處理
        print("✅ 修復完成")
        

    else:
        print("✅ 數據格式正常，無需修復。")
        # 先確認 header 是 1 還是 2，確保 columns 正確
        header = 2  # 預設標題行在第 3 行
        df = pd.read_csv(file_path, encoding=encoding, header=header)
    
        # 如果 "交易時間" 不在 columns，改用 header=1
        if "交易時間" not in df.columns:
            print("⚠️ 無法找到 '交易時間'，改用 header=1...")
            header = 1
            df = pd.read_csv(file_path, encoding=encoding, header=header)
    
        print(f"✅ 確定欄位標頭行數: header={header}")

    return df


# 使用函數
df = detect_and_fix_csv(input_file)
print(f"✅ 欄位名稱: {df.columns.tolist()}")

# 先保存原始交易時間用於調試
original_time_col = df["交易時間"].copy()

# 處理交易時間格式：統一轉換為標準格式
if df["交易時間"].dtype == 'object':
    # 轉換為字符串類型
    df["交易時間"] = df["交易時間"].astype(str)
    
    # 處理中文格式：將「上午」轉換為 " AM"，「下午」轉換為 " PM"
    df["交易時間"] = df["交易時間"].str.replace("上午", " AM", regex=False)
    df["交易時間"] = df["交易時間"].str.replace("下午", " PM", regex=False)
    
    # 處理小寫的 am/pm，轉換為大寫
    df["交易時間"] = df["交易時間"].str.replace(" am", " AM", regex=False)
    df["交易時間"] = df["交易時間"].str.replace(" pm", " PM", regex=False)
    
    # 確保英文格式的 AM/PM 有正確的空格（如果沒有空格，添加空格）
    # 例如 "2025/10/1 09:53:59AM" -> "2025/10/1 09:53:59 AM"
    df["交易時間"] = df["交易時間"].str.replace(r'(\d{1,2}:\d{2}:\d{2})(AM|PM)', r'\1 \2', regex=True)
    
    # 標準化多個空格為單個空格
    df["交易時間"] = df["交易時間"].str.replace(r'\s+', ' ', regex=True)
    
    # 移除首尾空格
    df["交易時間"] = df["交易時間"].str.strip()

# 嘗試多種格式解析交易時間
print("📅 開始解析交易時間...")

# 第一種格式：帶 AM/PM 的 12 小時制
df["交易時間"] = pd.to_datetime(df["交易時間"], format="%Y/%m/%d %p %I:%M:%S", errors="coerce")

# 如果第一種格式失敗，嘗試其他可能的格式
if df["交易時間"].isna().sum() > 0:
    invalid_count = df["交易時間"].isna().sum()
    print(f"⚠️ 有 {invalid_count} 筆資料無法用第一種格式解析，嘗試其他格式...")
    
    # 獲取無法解析的原始值（轉換前）
    mask = df["交易時間"].isna()
    invalid_original = original_time_col[mask].astype(str)
    
    # 嘗試不帶 AM/PM 的 24 小時制格式
    if mask.any():
        # 先嘗試 24 小時制
        df.loc[mask, "交易時間"] = pd.to_datetime(df.loc[mask, "交易時間"], format="%Y/%m/%d %H:%M:%S", errors="coerce")
        
        # 如果還有無法解析的，嘗試讓 pandas 自動推斷
        still_invalid = df["交易時間"].isna()
        if still_invalid.any():
            print(f"⚠️ 仍有 {still_invalid.sum()} 筆資料無法解析，嘗試自動推斷格式...")
            # 使用原始值重新嘗試解析
            df.loc[still_invalid, "交易時間"] = pd.to_datetime(original_time_col[still_invalid], errors="coerce")

# 檢查交易時間轉換是否成功
invalid_time_count = df["交易時間"].isna().sum()
if invalid_time_count > 0:
    print(f"⚠️ 警告：有 {invalid_time_count} 筆資料的交易時間無法解析")
    print(f"   前 5 筆無法解析的交易時間範例：")
    invalid_indices = df[df["交易時間"].isna()].index[:5]
    for idx in invalid_indices:
        print(f"     - {original_time_col.iloc[idx]}")

df["月份"] = df["交易時間"].dt.strftime("%Y-%m")

# 檢查月份欄位
valid_month_count = df["月份"].notna().sum()
print(f"✅ 有效月份資料筆數：{valid_month_count} / {len(df)}")
if valid_month_count == 0:
    print("❌ 錯誤：沒有有效的月份資料，無法產生檔案！")
    print("   請檢查交易時間欄位的格式")
else:
    print(f"   月份分布：{df['月份'].value_counts().to_dict()}")

# 初始化新欄位
df["油品類"] = None
df["油品類單價"] = None
df["油品類數量"] = None

# 定義油品類別對應關鍵字
# 注意：dict 順序就是匹配順序，loop 遇到第一個 match 就 break。
# 油料放前面 → 同群有油料時，油料優先成為油品類（添加劑被同群油料吸收）。
# 添加劑放後面 → 只有單獨買添加劑、沒搭油料的群組才會被分到添加劑類別。
fuel_mapping = {
    "九二": "九二無鉛汽油\(A02",
    "九五": "九五無鉛汽油",
    "超柴": "超級柴油",
    "尿素": "國光牌車用尿素",
    "柴油精": "柴油精",
    "汽油精": "汽油精",
    "機油精": "機油精",
}

# Per-row 標記 油品類（每筆商品列依其商品名稱判斷）
def classify_product(name):
    if not isinstance(name, str):
        return None
    for ftype, keyword in fuel_mapping.items():
        if re.search(keyword, name):
            return ftype
    return None

df["油品類"] = df["商品名稱"].apply(classify_product)

# 定義商品類別（付款摘要列；做為 summary 表的 row index）
categories = [
    "其他應收款-電子優惠券",
    "捷利卡優惠點消費－應稅",
    "捷利卡消費 (A53",
    "禮物卡消費－應稅",
    "車隊卡消費 (A55",
    "點數折抵主油A57"
]

# 把每個交易群組裡的「付款列金額」依「群內各商品金額」比例拆給每個商品
# 產生 (付款列商品名稱, 商品油品類) 的虛擬列，金額 = 商品金額 × (付款列金額/群內商品金額總和)
# 這樣油料/添加劑各歸各的欄，付款列總額（含多個折抵列）會正確分配。
virtual_rows = []
for _key, _g in df.groupby(["島號", "槍號", "交易時間"]):
    _products = _g[_g["油品類"].notna()]
    _payments = _g[_g["商品名稱"].isin(categories)]
    if _products.empty or _payments.empty:
        continue
    _sum_prod = _products["商品金額"].sum()
    if _sum_prod == 0:
        continue
    for _, _p in _payments.iterrows():
        # 保留正負號：付款金額/退款金額在後續 abs(sum()) 才能正確抵消
        _ratio = _p["商品金額"] / _sum_prod
        for _, _pr in _products.iterrows():
            virtual_rows.append({
                "商品名稱": _p["商品名稱"],
                "油品類": _pr["油品類"],
                "商品金額": _pr["商品金額"] * _ratio,
                "油品類數量": _pr["數量"] * _ratio,
                "月份": _g["月份"].iloc[0],
            })

if not virtual_rows:
    print("❌ 沒有可彙整的資料")
    input("\n按 Enter 鍵退出...")
    exit()

df_split = pd.DataFrame(virtual_rows)
print(f"\n📊 拆分後彙整列數：{len(df_split)}")

# ================================
# 發票與折扣表（每月一張，放在「油品交易彙總」檔的第二個頁籤）
# ================================
# 每筆交易（島號+槍號+交易時間）的「交易總金額」「發票金額」「折扣金額」是表頭欄位，群組內每列相同。
# 折扣額規則（已和手算班報逐日核對）：
#   - 發票金額 ≠ 0 且交易內沒有「點數折抵商品A58 / 點數折抵免稅商品A59」列 → 折扣額 = 表頭「折扣金額」欄
#   - 其他情況（車隊卡、捷利卡、禮物卡、簽帳等未開發票交易，以及點數兌換商品交易）→ 折扣額 = 0
#   對應班報：折讓 + 計程車優惠金額 + 計程車紅利金優惠 + 點數折抵主油。
#   點數折抵商品（A58/A59）的折抵金額另外列在「副產品 / 點數折抵額」，對應班報「點數兌換商品」。
# 一筆交易若含多個商品，表頭金額依各商品「商品金額」比例分攤到商品；
# 折扣額只分攤給油品（九二/九五/超柴/尿素），交易內沒有油品時才分攤給副產品。
# 每筆交易只歸到一個付款類別：取該筆交易內付款列金額（絕對值）最大的那個；
# 沒有任何付款列的交易（現金 / 刷卡 / 簽帳等）歸「其他(無付款列)」。
INVOICE_PRODUCTS = ["九二", "九五", "超柴", "尿素", "副產品"]
INVOICE_FUELS = ["九二", "九五", "超柴", "尿素"]
POINT_GOODS_ROWS = ["點數折抵商品A58", "點數折抵免稅商品A59"]
INVOICE_ROW_ORDER = categories + POINT_GOODS_ROWS
OTHER_ROW = "其他(現金/刷卡/簽帳，無付款列)"
TOTAL_ROW = "總計"
_all_payment_names = set(INVOICE_ROW_ORDER)


def invoice_product(name):
    """把商品列歸到 九二/九五/超柴/尿素/副產品；付款列與折抵額列回傳 None。"""
    if not isinstance(name, str):
        return None
    if name in _all_payment_names or "折抵額" in name:
        return None
    if "九二無鉛汽油" in name:
        return "九二"
    if "九五無鉛汽油" in name:
        return "九五"
    if "超級柴油" in name:
        return "超柴"
    if "尿素" in name:
        return "尿素"
    return "副產品"


def build_invoice_table(df_all):
    """回傳 (表格 DataFrame, 說明文字 list)。表格 index 為付款類別，欄為 (產品, 指標)。"""
    d = df_all[df_all["月份"].notna()].copy()
    d["_prod"] = d["商品名稱"].apply(invoice_product)
    d["_is_pay"] = d["商品名稱"].isin(_all_payment_names)
    d["_is_point_goods"] = d["商品名稱"].isin(POINT_GOODS_ROWS)
    key = ["島號", "槍號", "交易時間"]
    for c in ["交易總金額", "發票金額", "折扣金額", "商品金額"]:
        d[c] = pd.to_numeric(d[c], errors="coerce").fillna(0)

    # 每筆交易的主要付款類別：付款列金額絕對值最大者
    pay = d[d["_is_pay"]].copy()
    pay["_abs"] = pay["商品金額"].abs()
    pay_sum = pay.groupby(key + ["商品名稱"])["_abs"].sum().reset_index()
    pay_sum = pay_sum.sort_values("_abs", ascending=False).drop_duplicates(key)
    pay_sum = pay_sum.rename(columns={"商品名稱": "_row"})[key + ["_row"]]
    d = d.merge(pay_sum, on=key, how="left")
    d["_row"] = d["_row"].fillna(OTHER_ROW)

    # 折扣額：有發票且非點數兌換商品交易 → 表頭折扣金額；否則 0
    txn_has_point_goods = d.groupby(key)["_is_point_goods"].transform("any")
    d["_折扣"] = np.where((d["發票金額"] != 0) & (~txn_has_point_goods), d["折扣金額"], 0.0)

    # 表頭金額依商品金額比例分攤到各商品列
    prod = d[d["_prod"].notna()].copy()
    prod_sum = prod.groupby(key)["商品金額"].transform("sum")
    weight = np.where(prod_sum != 0, prod["商品金額"] / prod_sum.replace(0, np.nan), 0)
    prod["交易總金額_分攤"] = prod["交易總金額"] * weight
    prod["發票金額_分攤"] = prod["發票金額"] * weight
    # 折扣額優先分攤給油品；交易內沒有油品時才分攤給副產品
    is_fuel = prod["_prod"].isin(INVOICE_FUELS)
    fuel_sum = prod["商品金額"].where(is_fuel, 0).groupby([prod[k] for k in key]).transform("sum")
    has_fuel = prod["_prod"].isin(INVOICE_FUELS).groupby([prod[k] for k in key]).transform("any")
    disc_weight = np.where(
        has_fuel,
        np.where(is_fuel & (fuel_sum != 0), prod["商品金額"] / fuel_sum.replace(0, np.nan), 0),
        weight,
    )
    prod["折扣額_分攤"] = prod["_折扣"] * np.nan_to_num(disc_weight)

    agg = prod.groupby(["_row", "_prod"]).agg(
        交易總金額=("交易總金額_分攤", "sum"),
        發票金額=("發票金額_分攤", "sum"),
        折扣額=("折扣額_分攤", "sum"),
    )
    # 點數折抵商品（A58/A59）折抵金額，依付款類別列
    point_amt = -d[d["_is_point_goods"]].groupby("_row")["商品金額"].sum()

    rows = [r for r in INVOICE_ROW_ORDER if r in agg.index.get_level_values(0) or r in point_amt.index] + [OTHER_ROW]
    table = pd.DataFrame(index=pd.Index(rows, name="付款類別"))
    for p in INVOICE_PRODUCTS:
        metrics = ["發票金額", "點數折抵額"] if p == "副產品" else ["交易總金額", "發票金額", "折扣額"]
        for m in metrics:
            if m == "點數折抵額":
                table[(p, m)] = point_amt.reindex(rows).fillna(0).values
                continue
            vals = agg[m].unstack("_prod").get(p) if p in agg.index.get_level_values(1) else None
            table[(p, m)] = vals.reindex(rows).fillna(0).values if vals is not None else 0.0
    table[("合計", "發票金額")] = sum(table[(p, "發票金額")] for p in INVOICE_PRODUCTS)
    # 合計折扣額含分攤到副產品的極少數折扣（交易內沒有油品時），所以可能略大於四油品相加
    all_disc = agg["折扣額"].groupby("_row").sum().reindex(rows).fillna(0).values
    table[("合計", "折扣額")] = all_disc
    table.loc[TOTAL_ROW] = table.sum()
    table.columns = pd.MultiIndex.from_tuples(table.columns, names=["產品", "指標"])
    table = table.round(0).astype(int)

    notes = [
        "說明：",
        "1. 折扣額：發票金額≠0 且非點數兌換商品的交易，取表頭「折扣金額」欄；其他交易為 0。",
        "   對應班報：折讓 + 計程車優惠金額 + 計程車紅利金優惠 + 點數折抵主油。",
        "2. 車隊卡、捷利卡、禮物卡、簽帳交易現場不開發票，發票金額為 0，不算折扣。",
        "3. 副產品/點數折抵額 = 點數折抵商品A58、A59 的折抵金額，對應班報「點數兌換商品」。",
        "4. 一筆交易含多個商品時，表頭金額依各商品「商品金額」比例分攤；折扣額只分攤給油品。",
        "5. 每筆交易只歸到一個付款類別（付款列金額最大者）；無付款列者歸「其他」。",
        "6. 副產品 = 非九二/九五/超柴/尿素的其他商品。作廢交易為負值，已自動沖銷。",
    ]
    return table, notes


def month_label(month):
    """'2026-09' -> '2026年09月'，用在輸出檔名。"""
    y, m = str(month).split("-")
    return f"{y}年{m}月"


df_valid = df_split[df_split["月份"].notna()].copy()

# 逐月計算 summary 並存 Excel
file_count = 0
for month, df_month in df_valid.groupby("月份"):
    print(f"\n📅 處理月份：{month} (共 {len(df_month)} 筆拆分列)")

    df_filtered = df_month

    if len(df_filtered) == 0:
        print(f"   ⚠️ 該月份沒有符合商品類別的資料，跳過")
        continue

    try:
        summary_df = df_filtered.groupby(["商品名稱", "油品類"]).agg(
            油品類數量總=pd.NamedAgg(column="油品類數量", aggfunc=lambda x: abs(x.sum()) if x.notna().any() else 0),
            商品金額總=pd.NamedAgg(column="商品金額", aggfunc=lambda x: abs(x.sum()) if x.notna().any() else 0)
        ).unstack(fill_value=0)

        # 確保所有油品類（含添加劑）欄位都存在，缺的補 0
        _all_oil_types = ["九二", "九五", "尿素", "超柴", "柴油精", "汽油精", "機油精"]
        for _metric in ["油品類數量總", "商品金額總"]:
            for _oil in _all_oil_types:
                if (_metric, _oil) not in summary_df.columns:
                    summary_df[(_metric, _oil)] = 0
        # 重新排序欄位
        summary_df = summary_df.reindex(
            columns=pd.MultiIndex.from_product(
                [["油品類數量總", "商品金額總"], _all_oil_types],
                names=summary_df.columns.names,
            ),
            fill_value=0,
        )

        # 發票與折扣表
        invoice_table, invoice_notes = build_invoice_table(df[df["月份"] == month])

        # 存成 Excel：第一頁 油品彙總（原表），第二頁 發票與折扣
        output_file = os.path.join(current_dir, f"油品交易彙總_{month_label(month)}.xlsx")
        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            summary_df.to_excel(writer, sheet_name="油品彙總")
            invoice_table.to_excel(writer, sheet_name="發票與折扣")
            ws = writer.sheets["發票與折扣"]
            start = len(invoice_table) + 5
            for i, line in enumerate(invoice_notes):
                ws.cell(row=start + i, column=1, value=line)
            from openpyxl.utils import get_column_letter
            ws.column_dimensions["A"].width = 30
            for ci in range(2, ws.max_column + 1):
                ws.column_dimensions[get_column_letter(ci)].width = 12
        print(f"   ✅ 已儲存: {output_file}")
        file_count += 1
    except Exception as e:
        print(f"   ❌ 處理月份 {month} 時發生錯誤：{str(e)}")
        import traceback
        traceback.print_exc()

if file_count == 0:
    print("\n❌ 沒有產生任何檔案！")
else:
    print(f"\n✅ 成功產生 {file_count} 個月份的 summary 檔案！")

# ================================
# 作帳用：非汽油產品銷售明細
# ================================
# 排除汽油與超級柴油（含各自的折抵額列，名稱都含「無鉛汽油」或「超級柴油」），
# 以及付款/沖銷列（車隊卡消費、捷利卡消費、點數折抵等，非實際產品）。
print("\n================================")
print(" 產出作帳用非汽油產品明細")
print("================================")

payment_rows = categories + [
    "點數折抵商品A58",
    "點數折抵免稅商品A59",
]

# 點數折抵商品 / 非點數折抵商品：
#   以「交易時間 + 會員卡號」當一筆交易，只要該筆交易內出現
#   「點數折抵商品A58」或「點數折抵免稅商品A59」，該筆交易的所有商品都算點數折抵商品
#   （一筆交易可能有多個商品列）；否則為非點數折抵商品。
POINT_ROWS = ["點數折抵商品A58", "點數折抵免稅商品A59"]
SCOPE_ORDER = ["點數折抵商品", "非點數折抵商品"]

_txn_key = [
    df["交易時間"].astype(str),
    df["會員卡號"].fillna("").astype(str).str.strip() if "會員卡號" in df.columns else pd.Series("", index=df.index),
]
_is_point_row = df["商品名稱"].isin(POINT_ROWS)
df["商品範圍"] = np.where(
    _is_point_row.groupby(_txn_key).transform("any"),
    "點數折抵商品",
    "非點數折抵商品",
)

is_fuel = df["商品名稱"].astype(str).str.contains("無鉛汽油|超級柴油", na=False)
is_payment = df["商品名稱"].isin(payment_rows)
df_acct = df[~is_fuel & ~is_payment & df["月份"].notna()].copy()

acct_file_count = 0
if df_acct.empty:
    print("⚠️ 沒有非汽油產品的交易資料，跳過")
else:
    detail_cols = [
        "交易時間", "島號", "槍號", "發票號碼", "交易類別", "付款類別", "會員卡號", "商品範圍",
        "車號", "統一編號", "商品名稱", "單價", "數量", "商品金額",
    ]
    detail_cols = [c for c in detail_cols if c in df_acct.columns]

    for month, df_m in df_acct.groupby("月份"):
        print(f"\n📅 處理月份：{month} (共 {len(df_m)} 筆)")
        try:
            detail = df_m.sort_values("交易時間")[detail_cols].copy()
            detail["交易時間"] = detail["交易時間"].dt.strftime("%Y/%m/%d %H:%M:%S")

            # 彙總：依「商品範圍」把點數折抵商品 / 非點數折抵商品分開各自加總數量與金額
            # 金額為商品原價加總（退貨負數會正確抵消）
            by_scope = df_m.groupby(["商品名稱", "商品範圍"]).agg(
                數量=pd.NamedAgg(column="數量", aggfunc="sum"),
                金額=pd.NamedAgg(column="商品金額", aggfunc="sum"),
            ).unstack("商品範圍", fill_value=0)
            summary = pd.DataFrame(index=by_scope.index)
            for scope in SCOPE_ORDER:
                for metric in ["數量", "金額"]:
                    col = (metric, scope)
                    summary[f"{scope}_{metric}"] = by_scope[col] if col in by_scope.columns else 0
            # 依總金額由大到小排序（總金額只用來排序，不輸出）
            _order = (summary["點數折抵商品_金額"] + summary["非點數折抵商品_金額"]).sort_values(ascending=False).index
            summary = summary.loc[_order].reset_index()
            # 最後加一列「總計」
            total_row = {"商品名稱": "總計"}
            for c in summary.columns[1:]:
                total_row[c] = summary[c].sum()
            summary = pd.concat([summary, pd.DataFrame([total_row])], ignore_index=True)

            # 付款類別對照：列出每個付款類別代碼下點數折抵/非點數折抵的筆數與金額，方便核對
            pay_map_df = df_m.groupby(["付款類別", "商品範圍"]).agg(
                筆數=pd.NamedAgg(column="商品名稱", aggfunc="size"),
                金額=pd.NamedAgg(column="商品金額", aggfunc="sum"),
            ).reset_index()

            # 每個產品一個頁籤，頁籤名稱用產品名稱
            # Excel 頁籤名稱限制：最長 31 字元，不能含 / \ * ? [ ] : 等字元
            def safe_sheet_name(name, used):
                name = re.sub(r'[\\/*?\[\]:]', '_', str(name)).strip() or "未命名"
                name = name[:31]
                base = name
                n = 2
                while name in used:
                    suffix = f"_{n}"
                    name = base[:31 - len(suffix)] + suffix
                    n += 1
                used.add(name)
                return name

            output_file = os.path.join(current_dir, f"非油品作帳明細_{month_label(month)}.xlsx")
            with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
                sheets = [("彙總", summary), ("付款類別對照", pay_map_df)]
                used_names = {"彙總", "付款類別對照"}
                # 依彙總表的順序（銷售金額由大到小）建立各產品頁籤
                for product in summary["商品名稱"]:
                    if product == "總計":
                        continue
                    product_detail = detail[detail["商品名稱"] == product]
                    sheets.append((safe_sheet_name(product, used_names), product_detail))
                # 寫入並調整欄寬方便閱讀
                for sheet_name, frame in sheets:
                    frame.to_excel(writer, sheet_name=sheet_name, index=False)
                    ws = writer.sheets[sheet_name]
                    for i, col in enumerate(frame.columns, start=1):
                        max_len = max(
                            [len(str(col))] + [len(str(v)) for v in frame[col].head(200)]
                        )
                        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = min(max_len + 4, 40)
            print(f"   ✅ 已儲存: {output_file}")
            acct_file_count += 1
        except Exception as e:
            print(f"   ❌ 處理月份 {month} 時發生錯誤：{str(e)}")
            import traceback
            traceback.print_exc()

    if acct_file_count > 0:
        print(f"\n✅ 成功產生 {acct_file_count} 個月份的作帳用非汽油產品明細！")

input("\n按 Enter 鍵退出...")
