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
print("當前版本號 2026_1013_v7")

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

# 分月以「列帳班別」為準（格式 1150901_2 = 民國 115 年 09 月 01 日 第 2 班），
# 跟手算班報一致：跨午夜的晚班交易會算在列帳當天，而不是交易時間的日期。
# 無法解析列帳班別的列才退回用交易時間。
if "列帳班別" in df.columns:
    _lb = df["列帳班別"].astype(str).str.strip().str.extract(r"^(\d{3})(\d{2})(\d{2})_")
    _lb_ok = _lb[0].notna() & df["交易時間"].notna()
    df.loc[_lb_ok, "月份"] = (_lb.loc[_lb_ok, 0].astype(int) + 1911).astype(str) + "-" + _lb.loc[_lb_ok, 1]
    _moved = (_lb_ok & (df["月份"] != df["交易時間"].dt.strftime("%Y-%m"))).sum()
    print(f"📅 分月以列帳班別為準：{int(_lb_ok.sum())} / {len(df)} 列可解析，其中 {int(_moved)} 列的月份與交易時間不同")

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

# 定義商品類別（付款摘要列；做為 summary 表的 row index）
categories = [
    "其他應收款-電子優惠券",
    "捷利卡優惠點消費－應稅",
    "捷利卡消費 (A53",
    "禮物卡消費－應稅",
    "車隊卡消費 (A55",
    "點數折抵主油A57"
]
NON_PRODUCT_ROWS = set(categories + ["點數折抵商品A58", "點數折抵免稅商品A59"])
FIXED_PRODUCT_ORDER = list(fuel_mapping.keys())
OIL_SUMMARY_TYPES = ["九二", "九五", "超柴", "尿素"]  # 油品彙總表只輸出這四欄


# Per-row 標記 油品類（每筆商品列依其商品名稱判斷）
# 油品與添加劑用 fuel_mapping 的簡稱；其他商品（衛生紙、米、飲料…）直接用商品名稱，
# 這樣油品彙總表會列出所有商品，付款列的金額才能完整分攤。
def classify_product(name):
    if not isinstance(name, str):
        return None
    if name in NON_PRODUCT_ROWS or "折抵額" in name:
        return None
    for ftype, keyword in fuel_mapping.items():
        if re.search(keyword, name):
            return ftype
    return name.strip()

df["油品類"] = df["商品名稱"].apply(classify_product)

# 把每個交易群組裡的「付款列金額」依「群內各商品金額」比例拆給每個商品
# 產生 (付款列商品名稱, 商品油品類) 的虛擬列，金額 = 商品金額 × (付款列金額/群內商品金額總和)
# 這樣油料/添加劑各歸各的欄，付款列總額（含多個折抵列）會正確分配。
# 沒有付款列（或付款列金額不足整筆）的部分，依「付款類別」代碼另立一列，
# 這樣油品彙總的總計才會等於整月出油量。代碼對應是從班報金額表的合計反推：
#   901 現金、903 簽帳、905 車隊卡、931 信用卡、939 行動支付、900 試桶
TEST_BUCKET_PAY_TYPE = 900
PAY_TYPE_LABELS = {
    900: "試桶(付款類別900)",
    901: "現金(付款類別901)",
    903: "簽帳(付款類別903)",
    905: "車隊卡(付款類別905)",
    931: "信用卡(付款類別931)",
    939: "行動支付(付款類別939)",
}
POINT_GOODS_ROWS = ["點數折抵商品A58", "點數折抵免稅商品A59"]
_payment_names_for_split = categories + POINT_GOODS_ROWS


def pay_type_label(code):
    try:
        code = int(float(code))
    except (TypeError, ValueError):
        return "未知付款類別"
    return PAY_TYPE_LABELS.get(code, f"付款類別{code}")


virtual_rows = []
for _key, _g in df.groupby(["島號", "槍號", "交易時間"]):
    _products = _g[_g["油品類"].notna()]
    if _products.empty:
        continue
    _sum_prod = _products["商品金額"].sum()
    if _sum_prod == 0:
        continue
    _payments = _g[_g["商品名稱"].isin(_payment_names_for_split)]
    # 付款列金額為負；整筆商品金額扣掉付款列後剩下的，就是用現金/刷卡/簽帳等付的部分
    _pay_rows = [(_p["商品名稱"], _p["商品金額"]) for _, _p in _payments.iterrows()]
    _remain = _sum_prod + _payments["商品金額"].sum()
    if abs(_remain) >= 0.5:
        _pay_rows.append((pay_type_label(_g["付款類別"].iloc[0]), -_remain))
    for _name, _amt in _pay_rows:
        # 保留正負號：付款金額/退款金額在後續 abs(sum()) 才能正確抵消
        _ratio = _amt / _sum_prod
        for _, _pr in _products.iterrows():
            virtual_rows.append({
                "商品名稱": _name,
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
# 全部以「商品列」為基礎計算（已和手算班報「品項金額」頁逐日核對）：
#   總金額   = 各商品列「商品金額」加總（排除試桶 = 付款類別 900）；
#              副產品扣掉點數折抵商品A58/A59 的折抵金額，只留現金付的部分
#   折扣額   = 「92無鉛汽油折抵額」「95無鉛汽油折抵額」「超級柴油折抵額」等折抵額列，依油品歸類
#   農機免稅 = 有開發票但 交易總金額 − 發票金額 − 折扣金額 ≠ 0 的差額（付款類別 931、稅別 TF），
#              比照班報同時加進總金額與折扣額
#   折後金額 = 總金額 − 折扣額
# 捷利卡優惠點、紅利金、點數折抵主油等不算折扣，另外列在第二張表對班報「金額表」。
INVOICE_PRODUCTS = ["九二", "九五", "超柴", "尿素", "副產品"]
POINT_GOODS_ROWS = ["點數折抵商品A58", "點數折抵免稅商品A59"]
_all_payment_names = set(categories + POINT_GOODS_ROWS)


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


def discount_product(name):
    """折抵額列歸到哪個油品（名稱如 92無鉛汽油折抵額 / 95無鉛汽油折抵額 / 超級柴油免稅折抵額）。"""
    if "92" in name or "九二" in name:
        return "九二"
    if "95" in name or "九五" in name:
        return "九五"
    if "柴油" in name:
        return "超柴"
    if "尿素" in name:
        return "尿素"
    return "副產品"


def build_invoice_table(df_all):
    """回傳 (品項金額表, 其他項目表, 說明文字 list)。"""
    d = df_all[df_all["月份"].notna()].copy()
    key = ["島號", "槍號", "交易時間"]
    for c in ["交易總金額", "發票金額", "折扣金額", "商品金額"]:
        d[c] = pd.to_numeric(d[c], errors="coerce").fillna(0)
    d["_paytype"] = pd.to_numeric(d["付款類別"], errors="coerce")
    name = d["商品名稱"].astype(str)
    d["_prod"] = d["商品名稱"].apply(invoice_product)
    d["_is_discount_row"] = name.str.contains("折抵額", na=False)
    d["_is_test"] = d["_paytype"] == TEST_BUCKET_PAY_TYPE

    # ---- 總金額：商品列，排除試桶 ----
    prod = d[d["_prod"].notna() & ~d["_is_test"]]
    total = prod.groupby("_prod")["商品金額"].sum().reindex(INVOICE_PRODUCTS).fillna(0)
    point_goods = -d.loc[d["商品名稱"].isin(POINT_GOODS_ROWS), "商品金額"].sum()
    total["副產品"] -= point_goods

    # ---- 折扣額：折抵額列 ----
    dz = d[d["_is_discount_row"]].copy()
    dz["_dprod"] = dz["商品名稱"].astype(str).apply(discount_product)
    discount = (-dz.groupby("_dprod")["商品金額"].sum()).reindex(INVOICE_PRODUCTS).fillna(0)

    # ---- 農機免稅：有發票但 交易總 − 發票 − 折扣欄 ≠ 0，依交易內商品金額比例歸到商品 ----
    hdr = d.groupby(key).agg(交易總=("交易總金額", "first"), 發票=("發票金額", "first"), 折扣=("折扣金額", "first"))
    hdr["_exempt"] = np.where(hdr["發票"] != 0, hdr["交易總"] - hdr["發票"] - hdr["折扣"], 0.0)
    ex = prod.merge(hdr[["_exempt"]], left_on=key, right_index=True, how="left")
    ex_sum = ex.groupby(key)["商品金額"].transform("sum")
    ex["_share"] = ex["_exempt"].fillna(0) * np.where(ex_sum != 0, ex["商品金額"] / ex_sum.replace(0, np.nan), 0)
    exempt = ex.groupby("_prod")["_share"].sum().reindex(INVOICE_PRODUCTS).fillna(0)

    total_adj = total + exempt
    discount_adj = discount + exempt
    item = pd.DataFrame(
        [total_adj, discount_adj, total_adj - discount_adj],
        index=pd.Index(["總金額", "折扣額", "折後金額"], name="項目"),
    )
    item["合計"] = item.sum(axis=1)
    item = item.round(0).astype(int)

    # ---- 其他項目（對班報「金額表」）----
    def pay_sum(pname, mask=None):
        m = d["商品名稱"] == pname
        if mask is not None:
            m &= mask
        return -d.loc[m, "商品金額"].sum()

    has_bonus = d.groupby(key)["商品名稱"].transform(lambda s: (s == "捷利卡優惠點消費－應稅").any())
    jet_bonus = pay_sum("捷利卡消費 (A53", has_bonus)

    def split_payment(pname, mask=None):
        """把某付款列的金額依交易內各商品金額比例拆成 主油 / 尿素 / 副產品。"""
        m = d["商品名稱"] == pname
        if mask is not None:
            m &= mask
        pay_txn = -d.loc[m].groupby(key)["商品金額"].sum()
        pr = d[d["_prod"].notna() & ~d["_is_test"]].copy()
        pr["_grp"] = np.select([pr["_prod"].isin(["九二", "九五", "超柴"]), pr["_prod"] == "尿素"], ["主油", "尿素"], "副產品")
        pr = pr.merge(pay_txn.rename("_pay"), left_on=key, right_index=True, how="inner")
        txn_sum = pr.groupby(key)["商品金額"].transform("sum")
        pr["_alloc"] = pr["_pay"] * np.where(txn_sum != 0, pr["商品金額"] / txn_sum.replace(0, np.nan), 0)
        return pr.groupby("_grp")["_alloc"].sum().reindex(["主油", "尿素", "副產品"]).fillna(0)

    jet = split_payment("捷利卡消費 (A53", ~has_bonus)
    gift = split_payment("禮物卡消費－應稅")
    test_by_prod = d[d["_is_test"] & d["_prod"].notna()].groupby("_prod")["商品金額"].sum()
    invoices = hdr[hdr["發票"] != 0]
    others = pd.DataFrame(
        [
            ("車隊卡消費", pay_sum("車隊卡消費 (A55"), "班報：車隊卡(金額)"),
            ("捷利卡(主油)", jet["主油"], "班報：捷利卡(主油)"),
            ("捷利卡(尿素)", jet["尿素"], "班報：捷利卡(尿素)"),
            ("捷利卡(副產品)", jet["副產品"], "班報：捷利卡(副產品)"),
            ("捷利卡紅利金（與現金混合付款的捷利卡部分）", jet_bonus, "班報：計程車紅利金優惠"),
            ("捷利卡優惠點", pay_sum("捷利卡優惠點消費－應稅"), "班報：計程車優惠金額"),
            ("禮物卡(主油)", gift["主油"] + gift["尿素"], "班報：禮物卡(主油)"),
            ("禮物卡(副產品)", gift["副產品"], "班報：禮物卡(副產品)"),
            ("電子優惠券", pay_sum("其他應收款-電子優惠券"), "班報：電子優惠券(金額)"),
            ("點數折抵主油", pay_sum("點數折抵主油A57"), "班報：點數折抵主油(金額)"),
            ("點數折抵商品（A58+A59）", point_goods, "班報：點數兌換商品；已從副產品總金額扣除"),
            ("農機免稅", exempt.sum(), "班報：農機免稅；已加進總金額與折扣額"),
            ("試桶（付款類別 900）", test_by_prod.sum(), "班報：試桶；未計入總金額。" + "、".join(f"{k} {int(v)}" for k, v in test_by_prod.items())),
            ("發票金額（交易表頭合計）", hdr["發票"].sum(), "班報：當日發票金額合計"),
            ("發票張數", int((invoices["發票"] > 0).sum() - (invoices["發票"] < 0).sum()), "班報：發票數量（扣除作廢）"),
        ],
        columns=["項目", "金額", "對應班報"],
    )
    others["金額"] = others["金額"].round(0).astype(int)

    notes = [
        "說明：",
        "1. 總金額 = 各商品列「商品金額」加總，排除試桶（付款類別 900）；副產品已扣掉點數折抵商品 A58/A59 的折抵金額。",
        "2. 折扣額 = 92/95 無鉛汽油折抵額、超級柴油折抵額等折抵額列，依油品歸類。",
        "3. 農機免稅 = 有開發票但 交易總金額 − 發票金額 − 折扣金額 ≠ 0 的差額，比照班報同時加進總金額與折扣額。",
        "4. 折後金額 = 總金額 − 折扣額，對應班報「品項金額」頁的折後金額；不是 CSV 的發票金額欄（車隊卡、簽帳的發票金額為 0）。",
        "5. 捷利卡優惠點、紅利金、點數折抵主油不算折扣，列在第二張表對班報「金額表」。",
        "6. 日期以交易時間為準；班報以列帳班別為準，跨月晚班的交易會落在不同月份。作廢交易為負值，已自動沖銷。",
    ]
    return item, others, notes


# ================================
# 班報（格式比照手算班報 115.MM.xlsx：金額表 / 出油量表 / 品項金額）
# ================================
# 每班一欄：1 日 A 班、1 日 B 班、2 日 A 班 … 31 日 B 班，最後一欄月合計。
# A 班 = 列帳班別 _2，B 班 = 列帳班別 _3。
# CSV 算得出來的格子填數字，算不出來的（信用卡/行動支付各品牌、液面鏡、機械表、
# PD(四)、進油、噴油、農機優惠、二行程）留空；小計、月合計等維持班報原本的公式。
from openpyxl.styles import Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter

SHIFT_CODE_TO_AB = {"2": "A", "3": "B"}
DAYS_IN_SHEET = 31
FIRST_DATA_COL = 2                     # B
LAST_DATA_COL = FIRST_DATA_COL + DAYS_IN_SHEET * 2 - 1   # BK
TOTAL_COL = LAST_DATA_COL + 1          # BL


def shift_col(day, ab):
    return FIRST_DATA_COL + (int(day) - 1) * 2 + (0 if ab == "A" else 1)


def shift_series_to_cells(series):
    """series index = (day, 'A'/'B') → {col: value}"""
    out = {}
    for (day, ab), v in series.items():
        if pd.isna(day) or ab not in ("A", "B"):
            continue
        if 1 <= int(day) <= DAYS_IN_SHEET and v != 0:
            out[shift_col(day, ab)] = v
    return out


def build_shift_report_data(df_month):
    """算出班報三張表每一班的數字。回傳 dict: 表名 → {列號: {欄號: 值}}。"""
    d = df_month.copy()
    key = ["島號", "槍號", "交易時間"]
    for c in ["交易總金額", "發票金額", "折扣金額", "商品金額", "數量"]:
        d[c] = pd.to_numeric(d[c], errors="coerce").fillna(0)
    d["_paytype"] = pd.to_numeric(d["付款類別"], errors="coerce")
    ex = d["列帳班別"].astype(str).str.strip().str.extract(r"^\d{3}\d{2}(\d{2})_(\d)")
    d["_day"] = pd.to_numeric(ex[0], errors="coerce")
    d["_ab"] = ex[1].map(SHIFT_CODE_TO_AB)
    fallback = d["_day"].isna() | d["_ab"].isna()
    if fallback.any():
        d.loc[fallback, "_day"] = d.loc[fallback, "交易時間"].dt.day
        d.loc[fallback, "_ab"] = "A"
    d["_day"] = d["_day"].astype(int)
    name = d["商品名稱"].astype(str)
    d["_prod"] = d["商品名稱"].apply(invoice_product)
    d["_is_disc"] = name.str.contains("折抵額", na=False)
    d["_is_test"] = d["_paytype"] == TEST_BUCKET_PAY_TYPE
    d["_grp"] = np.select([d["_prod"].isin(["九二", "九五", "超柴"]), d["_prod"] == "尿素"], ["主油", "尿素"], "副產品")
    d.loc[d["_prod"].isna(), "_grp"] = None

    g = d.groupby(key)
    txn = g.agg(
        day=("_day", "first"), ab=("_ab", "first"), pc=("_paytype", "first"),
        交易總=("交易總金額", "first"), 發票=("發票金額", "first"), 折扣=("折扣金額", "first"),
        發票號=("發票號碼", "first"), 作廢=("是否作廢", "first"),
    )
    txn["has_bonus"] = g["商品名稱"].agg(lambda s: (s == "捷利卡優惠點消費－應稅").any())
    txn["has_point"] = g["商品名稱"].agg(lambda s: s.isin(POINT_GOODS_ROWS).any())
    txn["exempt"] = np.where(txn["發票"] != 0, txn["交易總"] - txn["發票"] - txn["折扣"], 0.0)
    d = d.merge(txn[["pc", "發票", "has_bonus", "has_point", "exempt"]].rename(columns={"pc": "_pc", "發票": "_inv"}),
                left_on=key, right_index=True, how="left")
    sk = ["_day", "_ab"]

    def by_shift(frame, col):
        if frame.empty:
            return pd.Series(dtype=float)
        s = frame.groupby(sk)[col].sum()
        s.index = pd.MultiIndex.from_tuples([(int(a), b) for a, b in s.index])
        return s

    prod = d[d["_prod"].notna() & ~d["_is_test"]]
    is_sign = (d["_pc"] == 903) & (d["_inv"] == 0)
    is_fleet = d["_pc"] == 905

    # 付款列依交易內商品金額比例拆成 主油/尿素/副產品
    def split_payment(pname, txn_mask):
        m = (d["商品名稱"] == pname) & txn_mask
        pay = -d.loc[m].groupby(key)["商品金額"].sum()
        pr = d[d["_prod"].notna() & ~d["_is_test"]].merge(pay.rename("_pay"), left_on=key, right_index=True, how="inner")
        tsum = pr.groupby(key)["商品金額"].transform("sum")
        pr["_alloc"] = pr["_pay"] * np.where(tsum != 0, pr["商品金額"] / tsum.replace(0, np.nan), 0)
        return {grp: by_shift(pr[pr["_grp"] == grp], "_alloc") for grp in ["主油", "尿素", "副產品"]}

    jet = split_payment("捷利卡消費 (A53", ~d["has_bonus"].fillna(False))
    gift = split_payment("禮物卡消費－應稅", pd.Series(True, index=d.index))

    def pay_rows(pname, txn_mask=None):
        m = d["商品名稱"] == pname
        if txn_mask is not None:
            m &= txn_mask
        return -by_shift(d[m], "商品金額")

    # 農機免稅依交易內油品金額比例分到油品
    exempt_txn = txn[txn["exempt"] != 0]["exempt"]
    if not exempt_txn.empty:
        pe = prod.merge(exempt_txn.rename("_ex"), left_on=key, right_index=True, how="inner")
        tsum = pe.groupby(key)["商品金額"].transform("sum")
        pe["_exs"] = pe["_ex"] * np.where(tsum != 0, pe["商品金額"] / tsum.replace(0, np.nan), 0)
    else:
        pe = prod.iloc[0:0].assign(_exs=[])

    def exempt_of(p):
        return by_shift(pe[pe["_prod"] == p], "_exs") if not pe.empty else pd.Series(dtype=float)

    disc = d[d["_is_disc"]].copy()
    disc["_dprod"] = disc["商品名稱"].astype(str).apply(discount_product)

    def disc_of(p):
        return -by_shift(disc[disc["_dprod"] == p], "商品金額")

    def prod_amt(p, mask=None):
        f = prod[prod["_prod"] == p] if mask is None else prod[(prod["_prod"] == p) & mask.reindex(prod.index).fillna(False)]
        return by_shift(f, "商品金額")

    def add(*series):
        out = None
        for s_ in series:
            if s_ is None or len(s_) == 0:
                continue
            s_ = s_.copy()
            s_.index = pd.MultiIndex.from_tuples(list(s_.index))
            out = s_ if out is None else out.add(s_, fill_value=0)
        return out if out is not None else pd.Series(dtype=float)

    point_goods_amt = -by_shift(d[d["商品名稱"].isin(POINT_GOODS_ROWS)], "商品金額")
    point_txn_goods = prod_amt("副產品", d["has_point"].fillna(False))

    尿素_total = prod_amt("尿素")
    尿素_簽帳 = prod_amt("尿素", is_sign)
    尿素_車隊卡 = prod_amt("尿素", is_fleet)
    尿素_捷利卡 = jet["尿素"]
    尿素_現銷 = add(尿素_total, -尿素_簽帳, -尿素_車隊卡, -尿素_捷利卡)

    副_total = prod_amt("副產品")
    副_簽帳 = prod_amt("副產品", is_sign)
    副_捷利卡 = jet["副產品"]
    副_兌換 = point_txn_goods
    副_現銷 = add(副_total, -副_簽帳, -副_捷利卡, -副_兌換)

    txn_inv = txn[txn["發票號"].notna() & (txn["作廢"] != 1)]
    inv_cnt = txn_inv.groupby(["day", "ab"])["發票號"].nunique() if not txn_inv.empty else pd.Series(dtype=float)
    inv_cnt.index = pd.MultiIndex.from_tuples([(int(a), b) for a, b in inv_cnt.index]) if len(inv_cnt) else inv_cnt.index

    def txn_sum(mask, col):
        t = txn[mask]
        if t.empty:
            return pd.Series(dtype=float)
        s = t.groupby(["day", "ab"])[col].sum()
        s.index = pd.MultiIndex.from_tuples([(int(a), b) for a, b in s.index])
        return s

    金額表 = {
        4: txn_sum((txn["pc"] == 903) & (txn["發票"] == 0), "交易總"),           # 簽帳
        9: txn_sum(txn["pc"] == 931, "發票"),                                  # 信用卡小計
        22: txn_sum(txn["pc"] == 939, "發票"),                                 # 行動支付
        24: pay_rows("車隊卡消費 (A55"),
        26: jet["主油"], 27: jet["尿素"], 28: jet["副產品"],
        29: add(gift["主油"], gift["尿素"]), 30: gift["副產品"],
        31: pay_rows("捷利卡優惠點消費－應稅"),
        32: pay_rows("捷利卡消費 (A53", d["has_bonus"].fillna(False)),
        35: pay_rows("其他應收款-電子優惠券"),
        37: pay_rows("點數折抵主油A57"),
        39: 尿素_現銷, 40: 尿素_簽帳, 41: 尿素_車隊卡, 42: 尿素_捷利卡,
        46: 副_簽帳, 47: 副_捷利卡, 48: 副_現銷, 49: 副_兌換,
        52: point_goods_amt,
        54: txn_sum(txn["exempt"] != 0, "exempt"),
        56: -by_shift(disc, "商品金額"),
        59: by_shift(d[d["_is_test"] & d["_prod"].notna()], "商品金額"),
        64: inv_cnt,
    }
    品項金額 = {
        3: add(prod_amt("九二"), exempt_of("九二")),
        4: add(prod_amt("九五"), exempt_of("九五")),
        5: add(prod_amt("超柴"), exempt_of("超柴")),
        6: 尿素_total,
        7: add(副_total, -point_goods_amt),
        9: add(disc_of("九二"), exempt_of("九二")),
        10: add(disc_of("九五"), exempt_of("九五")),
        11: add(disc_of("超柴"), exempt_of("超柴")),
        12: add(disc_of("尿素"), exempt_of("尿素")),
        13: disc_of("副產品"),
    }
    qty = lambda p: by_shift(prod[prod["_prod"] == p], "數量")
    出油量表 = {
        3: qty("九二").round(0), 4: qty("九五").round(0), 5: qty("超柴").round(0),
        22: qty("尿素").round(2),
    }
    to_cells = lambda table: {r: shift_series_to_cells(s_.round(2)) for r, s_ in table.items()}
    return {"金額表": to_cells(金額表), "出油量表": to_cells(出油量表), "品項金額": to_cells(品項金額)}


# 班報版面：每張表的列標籤與公式（{c} 會換成該班的欄字母，{p} 為前一欄、{t} 為月合計欄）
_FONT = "新細明體"
_LABEL_FONT_SIZE = {"金額表": 12, "出油量表": 14, "品項金額": 14}

SHIFT_SHEETS = {
    "金額表": {
        "labels": {3: "現金", 4: "簽帳", 6: "中國信託", 7: "聯合信用卡", 8: "美國運通", 9: "信用卡小計",
                   11: "台灣PAY", 12: "中油PAY(CPC)", 13: "LINE PAY", 14: "Pi錢包", 15: "一卡通(ipass)",
                   16: "街口(jko pay)", 17: "歐付寶", 18: "橘子支付", 19: "車麻吉", 20: "iCash", 21: "悠遊付",
                   22: "行動支付", 24: "車隊卡(金額)", 26: "捷利卡(主油)", 27: "捷利卡(尿素)", 28: "捷利卡(副產品)",
                   29: "禮物卡(主油)", 30: "禮物卡(副產品)", 31: "計程車優惠金額", 32: "計程車紅利金優惠", 33: "小計",
                   35: "電子優惠券(金額)", 37: "點數折抵主油(金額)", 39: "尿素(現銷)", 40: "尿素(簽帳)",
                   41: "尿素(車隊卡)", 42: "尿素(捷利卡)", 43: "小計", 45: "二行程", 46: "副產品(簽帳)",
                   47: "副產品(捷利卡)", 48: "副產品(現銷)", 49: "副產品(兌換)", 50: "小計", 52: "點數兌換商品",
                   54: "農機免稅", 55: "農機優惠", 56: "折讓", 57: "小計", 59: "試桶", 60: "噴油",
                   61: "主油銷售額", 62: "實際營業額收入", 64: "發票數量", 65: "發票數量小計", 66: "當日發票金額"},
        "shift_formulas": {3: "=SUM({c}66-{c}9-{c}22)", 33: "=SUM({c}26:{c}32)", 43: "=SUM({c}39:{c}42)",
                           50: "=SUM({c}45:{c}49)", 57: "=SUM({c}54:{c}56)",
                           61: "=SUM(品項金額!{c}3:{c}5)",
                           62: "=SUM({c}61-{c}57+{c}43+{c}50-{c}52)",
                           66: "=SUM({c}62-{c}24-{c}33-{c}35-{c}37-{c}4)"},
        "day_formulas": {65: "=SUM({c}64:{n}64)"},   # 只在 A 班欄，{n} = B 班欄
        "total_rows": [3, 4, 6, 7, 8, 9, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 24, 26, 27, 28, 29, 30, 31, 32,
                       35, 37, 39, 40, 41, 42, 45, 46, 47, 48, 49, 52, 54, 55, 56, 59, 61, 64, 65, 66],
        "total_formulas": {22: "=SUM({t}11:{t}21)", 33: "=SUM({t}26:{t}32)", 43: "=SUM({t}39:{t}42)",
                           50: "=SUM({t}45:{t}49)", 57: "=SUM({t}54:{t}56)",
                           62: "=SUM({t}61-{t}57+{t}43+{t}50-{t}52)"},
        "bold_label_rows": None,
        "width_a": 22.9, "width_data": 9.4, "last_row": 68,
    },
    "出油量表": {
        "labels": {3: "92", 4: "95", 5: "PD(一)", 6: "PD(四)", 7: "班小計", 8: "日總計(電腦帳)",
                   10: "92", 11: "95", 12: "PD(一)", 13: "PD(四)", 14: "班小計", 15: "日總計(液面鏡)",
                   17: "92", 18: "95", 19: "PD", 20: "日總計(機械表)",
                   22: "尿素", 23: "日總計(電腦帳)", 25: "尿素", 26: "日總計(機械表)",
                   28: "92", 29: "95", 30: "PD(一)", 31: "PD(四)", 32: "班小計"},
        "shift_formulas": {7: "=SUM({c}3:{c}6)", 14: "=SUM({c}10:{c}13)", 20: "=SUM({c}17:{c}19)",
                           28: "=SUM({c}10-{c}3)", 29: "=SUM({c}11-{c}4)", 30: "=SUM({c}12-{c}5)",
                           31: "=SUM({c}13-{c}6)", 32: "=SUM({c}28:{c}31)"},
        "day_formulas": {8: "={c}7+{n}7", 15: "={c}14+{n}14", 23: "=SUM({c}22+{n}22)", 26: "=SUM({c}25+{n}25)"},
        "total_rows": [3, 4, 5, 6, 8, 10, 11, 12, 13, 15, 17, 18, 19, 20, 22, 23, 25, 26],
        "total_formulas": {7: "=SUM({t}3:{t}6)", 14: "=SUM({t}10:{t}13)"},
        "bold_label_rows": None,
        "width_a": 29.0, "width_data": 11.8, "last_row": 34,
    },
    "品項金額": {
        "labels": {3: "92", 4: "95", 5: "超柴", 6: "尿素", 7: "副產品", 9: "92折", 10: "95折", 11: "超柴折",
                   12: "尿素折", 13: "副產品", 17: "折後金額", 18: "92", 19: "95", 20: "超柴", 21: "尿素",
                   22: "副產品", 23: "銷貨金額"},
        "shift_formulas": {18: "=SUM({c}3-{c}9)", 19: "=SUM({c}4-{c}10)", 20: "=SUM({c}5-{c}11)",
                           21: "=SUM({c}6-{c}12)", 22: "=SUM({c}7-{c}13)", 23: "=SUM({c}18:{c}22)"},
        "day_formulas": {},
        "total_rows": [3, 4, 5, 6, 7, 9, 10, 11, 12, 13, 18, 19, 20, 21, 22, 23],
        "total_formulas": {},
        "bold_label_rows": None,
        "width_a": 12.4, "width_data": 9.4, "last_row": 23,
    },
}


def write_shift_report(output_file, data, month):
    """依班報版面把 data（build_shift_report_data 的結果）寫成 xlsx。"""
    import openpyxl
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    thin, medium = Side(style="thin"), Side(style="medium")
    center = Alignment(horizontal="center", vertical="center")
    y, m = str(month).split("-")
    roc_label = f"{int(y) - 1911}.{m}"

    for sheet_name, spec in SHIFT_SHEETS.items():
        ws = wb.create_sheet(sheet_name)
        size = _LABEL_FONT_SIZE[sheet_name]
        last_row = spec["last_row"]
        tcol = get_column_letter(TOTAL_COL)
        # 表頭
        ws.cell(1, 1, "日期")
        ws.cell(2, 1, "班別")
        for day in range(1, DAYS_IN_SHEET + 1):
            ca = shift_col(day, "A")
            ws.cell(1, ca, day)
            ws.merge_cells(start_row=1, start_column=ca, end_row=1, end_column=ca + 1)
            ws.cell(2, ca, "A")
            ws.cell(2, ca + 1, "B")
        ws.cell(1, TOTAL_COL, "月合計")
        # 標籤
        for r, label in spec["labels"].items():
            ws.cell(r, 1, label)
        # 數字
        for r, cells in data.get(sheet_name, {}).items():
            for c, v in cells.items():
                ws.cell(r, c, float(v) if abs(v - round(v)) > 1e-9 else int(round(v)))
        # 每班公式
        for day in range(1, DAYS_IN_SHEET + 1):
            for ab in ("A", "B"):
                c = get_column_letter(shift_col(day, ab))
                for r, f in spec["shift_formulas"].items():
                    ws.cell(r, shift_col(day, ab), f.format(c=c, t=tcol))
            ca = shift_col(day, "A")
            for r, f in spec["day_formulas"].items():
                ws.cell(r, ca, f.format(c=get_column_letter(ca), n=get_column_letter(ca + 1), t=tcol))
        # 月合計
        first, last = get_column_letter(FIRST_DATA_COL), get_column_letter(LAST_DATA_COL)
        for r in spec["total_rows"]:
            ws.cell(r, TOTAL_COL, f"=SUM({first}{r}:{last}{r})")
        for r, f in spec["total_formulas"].items():
            ws.cell(r, TOTAL_COL, f.format(t=tcol))
        # 樣式
        for r in range(1, last_row + 1):
            for c in range(1, TOTAL_COL + 1):
                cell = ws.cell(r, c)
                is_label = c == 1
                cell.font = Font(name=_FONT, size=14 if r <= 2 else size if is_label else 12, bold=is_label or r <= 2)
                cell.alignment = center
                cell.border = Border(
                    left=medium if c == 1 else thin, right=medium if c == TOTAL_COL else thin,
                    top=medium if r == 1 else thin, bottom=medium if r == last_row else thin,
                )
        ws.column_dimensions["A"].width = spec["width_a"]
        for c in range(FIRST_DATA_COL, TOTAL_COL + 1):
            ws.column_dimensions[get_column_letter(c)].width = spec["width_data"] if c < TOTAL_COL else 11
        ws.row_dimensions[1].height = 19.8
        ws.row_dimensions[2].height = 19.8
        ws.freeze_panes = "B3"

    ws = wb["金額表"]
    ws.cell(68, 1, roc_label).font = Font(name=_FONT, size=12, bold=True)
    # 品項金額右側的備註欄（比照班報）
    ws = wb["品項金額"]
    ws.cell(1, TOTAL_COL + 1, "車隊卡折讓")
    ws.cell(1, TOTAL_COL + 2, "賖銷折讓")
    # 出油量表下方的進油表（只留表頭，進油量 CSV 沒有）
    ws = wb["出油量表"]
    for c, v in zip(range(2, 6), ["日期", "92", "95", "PD(一)"]):
        ws.cell(34, c, v).font = Font(name=_FONT, size=12)
    # 說明
    notes = [
        f"由 交易查詢會總表.csv 產生（{month}）。A 班 = 列帳班別 _2，B 班 = 列帳班別 _3。",
        "算不出來的格子留空：信用卡與行動支付各品牌（只填小計）、液面鏡、機械表、PD(四)（超柴全部放 PD(一)）、進油、噴油、農機優惠、二行程。",
        "現金 = 當日發票金額 − 信用卡 − 行動支付；信用卡 = 付款類別 931 發票金額；行動支付 = 939 發票金額；簽帳 = 903 未開發票交易總金額。",
        "折讓 = 折抵額列；農機免稅 = 有發票但 交易總 − 發票 − 折扣金額 ≠ 0 的差額（同時加進品項金額的原價與折）；試桶 = 付款類別 900。",
    ]
    ws = wb["金額表"]
    for i, line in enumerate(notes):
        ws.cell(70 + i, 1, line).font = Font(name=_FONT, size=10)
    wb.save(output_file)


def month_label(month):
    """'2026-09' -> '2026年09月'，用在輸出檔名。"""
    y, m = str(month).split("-")
    return f"{y}年{m}月"


df_valid = df_split[df_split["月份"].notna()].copy()

# 逐月計算 summary 並存 Excel
file_count = 0
for month, df_month in df_valid.groupby("月份"):
    print(f"\n📅 處理月份：{month} (共 {len(df_month)} 筆拆分列)")

    # 油品彙總只列四種油品；其他商品仍參與付款列的比例拆分，只是不輸出
    df_filtered = df_month[df_month["油品類"].isin(OIL_SUMMARY_TYPES)]

    if len(df_filtered) == 0:
        print(f"   ⚠️ 該月份沒有符合商品類別的資料，跳過")
        continue

    try:
        summary_df = df_filtered.groupby(["商品名稱", "油品類"]).agg(
            油品類數量總=pd.NamedAgg(column="油品類數量", aggfunc=lambda x: abs(x.sum()) if x.notna().any() else 0),
            商品金額總=pd.NamedAgg(column="商品金額", aggfunc=lambda x: abs(x.sum()) if x.notna().any() else 0)
        ).unstack(fill_value=0)

        # 欄位固定為四種油品，缺的補 0
        _all_oil_types = OIL_SUMMARY_TYPES
        summary_df = summary_df.reindex(
            columns=pd.MultiIndex.from_product(
                [["油品類數量總", "商品金額總"], _all_oil_types],
                names=summary_df.columns.names,
            ),
            fill_value=0,
        )
        # 金額合計欄（數量單位不同，不加總）
        summary_df[("商品金額總", "合計")] = summary_df["商品金額總"].sum(axis=1)
        # 列順序：付款列類別 → 各付款類別代碼 → 總計；試桶另外放在總計下方，不計入總計
        _test_label = PAY_TYPE_LABELS[TEST_BUCKET_PAY_TYPE]
        _row_order = categories + POINT_GOODS_ROWS + [v for v in PAY_TYPE_LABELS.values() if v != _test_label]
        _rows = [c for c in _row_order if c in summary_df.index] + sorted(
            c for c in summary_df.index if c not in _row_order and c != _test_label
        )
        _test_row = summary_df.loc[[_test_label]] if _test_label in summary_df.index else None
        summary_df = summary_df.reindex(index=_rows)
        summary_df.loc["總計"] = summary_df.sum()
        if _test_row is not None:
            summary_df = pd.concat([summary_df, _test_row])

        # 發票與折扣表
        item_table, other_table, invoice_notes = build_invoice_table(df[df["月份"] == month])

        # 存成 Excel：第一頁 油品彙總（原表），第二頁 發票與折扣
        output_file = os.path.join(current_dir, f"油品交易彙總_{month_label(month)}.xlsx")
        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            summary_df.to_excel(writer, sheet_name="油品彙總")
            sheet = "發票與折扣"
            item_table.to_excel(writer, sheet_name=sheet, startrow=1)
            ws = writer.sheets[sheet]
            ws.cell(row=1, column=1, value="品項金額（對班報「品項金額」頁）")
            row2 = len(item_table) + 5
            ws.cell(row=row2 - 1, column=1, value="其他項目（對班報「金額表」）")
            other_table.to_excel(writer, sheet_name=sheet, startrow=row2 - 1, index=False)
            row3 = row2 + len(other_table) + 3
            for i, line in enumerate(invoice_notes):
                ws.cell(row=row3 + i, column=1, value=line)
            from openpyxl.utils import get_column_letter
            ws.column_dimensions["A"].width = 42
            ws.column_dimensions["B"].width = 14
            ws.column_dimensions["C"].width = 60
            for ci in range(4, ws.max_column + 1):
                ws.column_dimensions[get_column_letter(ci)].width = 12
        print(f"   ✅ 已儲存: {output_file}")
        file_count += 1

        # 班報（比照手算班報格式）
        report_file = os.path.join(current_dir, f"班報_{month_label(month)}.xlsx")
        write_shift_report(report_file, build_shift_report_data(df[df["月份"] == month]), month)
        print(f"   ✅ 已儲存: {report_file}")
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
