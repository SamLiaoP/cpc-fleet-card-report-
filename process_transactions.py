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

# 過濾掉月份為 NaN 的資料
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

        # 存成 Excel
        output_file = os.path.join(current_dir, f"交易_summary_{month}.xlsx")
        summary_df.to_excel(output_file)
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
input("\n按 Enter 鍵退出...")
