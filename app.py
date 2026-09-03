import streamlit as st
import pandas as pd
import unicodedata

st.set_page_config(page_title="売上明細 照合システム", layout="wide")

st.title("🚗 売上明細 照合・差分抽出システム")
st.caption("`deals.csv` と `整備売上明細.csv` を照合し、不一致（片方にしか存在しない）データを抽出します。")

# --- サイドバー：設定・除外ルール ---
st.sidebar.header("⚙️ 照合ルール設定")

valid_kanjo_default = "整備, 点検整備, 車検, 鈑金, レンタル売上"
valid_kanjo_input = st.sidebar.text_input("対象とする勘定科目 (カンマ区切り)", valid_kanjo_default)
valid_kanjo = [x.strip() for x in valid_kanjo_input.split(",") if x.strip()]

st.sidebar.markdown("---")
st.sidebar.subheader("除外ルール設定")
ex_tenken_pack = st.sidebar.checkbox("「点検整備」かつ「定期点検パック(オリジナル)」を除外", value=True)
ex_denpyo_nashi = st.sidebar.checkbox("品目「伝票なし」を除外", value=True)
ex_leaseup = st.sidebar.checkbox("「リースアップ」を含むデータを除外", value=True)

ex_seibi_kinds_input = st.sidebar.text_input("整備側で除外する種別 (カンマ区切り)", "その他, 新車待ち代車")
ex_seibi_kinds = [x.strip() for x in ex_seibi_kinds_input.split(",") if x.strip()]

# --- ヘルパー関数 ---
def read_csv_safe(file):
    """文字コード自動判定付きCSV読み込み"""
    try:
        return pd.read_csv(file, encoding='cp932')
    except Exception:
        file.seek(0)
        return pd.read_csv(file, encoding='utf-8-sig')

def clean_reg(text):
    if pd.isna(text):
        return ""
    text = unicodedata.normalize('NFKC', str(text))
    return text.replace(' ', '').replace(' ', '').replace('-', '')

def normalize_type_deals(row):
    pimmoku = row.get('品目', '')
    kanjo = row.get('勘定科目', '')
    if pd.isna(pimmoku) or str(pimmoku).strip() == '':
        if kanjo == '車検':
            return '車検'
        return ""
    val = unicodedata.normalize('NFKC', str(pimmoku)).strip()
    if val == '12ヵ月点検':
        return '12ヶ月点検'
    if val == '納車前点検':
        return '納車'
    return val

def normalize_type_seibi(val):
    if pd.isna(val):
        return ""
    val = unicodedata.normalize('NFKC', str(val)).strip()
    if val in ['整備', 'タイヤ交換']:
        return '一般整備'
    if val == '12ヵ月点検':
        return '12ヶ月点検'
    return val

# --- メイン画面：ファイルアップロード ---
col1, col2 = st.columns(2)

with col1:
    file_deals = st.file_uploader("1. deals.csv をアップロード", type=["csv"])

with col2:
    file_seibi = st.file_uploader("2. 整備売上明細 CSV をアップロード", type=["csv"])

if file_deals and file_seibi:
    try:
        df_deals_orig = read_csv_safe(file_deals)
        df_seibi_orig = read_csv_safe(file_seibi)

        # 必須列チェック
        req_deals_cols = ['備考', '勘定科目', '品目', '金額', '税額', '管理番号', '取引先', '発生日']
        req_seibi_cols = ['登録番号', '種別', '請求額（税抜）', '伝票No', '請求先名', '売上日']

        missing_deals = [c for c in req_deals_cols if c not in df_deals_orig.columns]
        missing_seibi = [c for c in req_seibi_cols if c not in df_seibi_orig.columns]

        if missing_deals:
            st.error(f"❌ deals.csv に必要な列が見つかりません: {', '.join(missing_deals)}")
            st.stop()
        if missing_seibi:
            st.error(f"❌ 整備売上明細CSV に必要な列が見つかりません: {', '.join(missing_seibi)}")
            st.stop()

        # --- データ前処理 & フィルタリング ---
        # deals 側フィルタ
        df_deals = df_deals_orig[df_deals_orig['勘定科目'].isin(valid_kanjo)].copy()

        exclude_deals_conditions = []
        if ex_tenken_pack:
            exclude_deals_conditions.append((df_deals['勘定科目'] == '点検整備') & (df_deals['品目'] == '定期点検パック(オリジナル)'))
        if ex_denpyo_nashi:
            exclude_deals_conditions.append(df_deals['品目'] == '伝票なし')
        if ex_leaseup:
            exclude_deals_conditions.append(df_deals['品目'].astype(str).str.contains('リースアップ', na=False))

        if exclude_deals_conditions:
            final_deals_mask = exclude_deals_conditions[0]
            for cond in exclude_deals_conditions[1:]:
                final_deals_mask = final_deals_mask | cond
            df_deals = df_deals[~final_deals_mask].copy()

        # seibi 側フィルタ
        df_seibi = df_seibi_orig.copy()
        exclude_seibi_conditions = []
        if ex_seibi_kinds:
            exclude_seibi_conditions.append(df_seibi['種別'].isin(ex_seibi_kinds))
        if ex_leaseup:
            exclude_seibi_conditions.append(df_seibi['種別'].astype(str).str.contains('リースアップ', na=False))

        if exclude_seibi_conditions:
            final_seibi_mask = exclude_seibi_conditions[0]
            for cond in exclude_seibi_conditions[1:]:
                final_seibi_mask = final_seibi_mask | cond
            df_seibi = df_seibi[~final_seibi_mask].copy()

        # --- 照合キー作成 ---
        df_deals['orig_index'] = df_deals.index
        df_seibi['orig_index'] = df_seibi.index

        df_deals['reg_clean'] = df_deals['備考'].apply(clean_reg)
        df_seibi['reg_clean'] = df_seibi['登録番号'].apply(clean_reg)

        df_deals['type_clean'] = df_deals.apply(normalize_type_deals, axis=1)
        df_seibi['type_clean'] = df_seibi['種別'].apply(normalize_type_seibi)

        df_deals['amount_clean'] = (df_deals['金額'] - df_deals['税額'].fillna(0)).round().astype(int)
        df_seibi['amount_clean'] = pd.to_numeric(df_seibi['請求額（税抜）'], errors='coerce').fillna(0).round().astype(int)

        df_deals['match_key'] = df_deals['type_clean'] + "_" + df_deals['reg_clean'] + "_" + df_deals['amount_clean'].astype(str)
        df_seibi['match_key'] = df_seibi['type_clean'] + "_" + df_seibi['reg_clean'] + "_" + df_seibi['amount_clean'].astype(str)

        # --- 照合ループ実行 ---
        seibi_matched_indices = set()
        deals_unmatched = []

        for idx, row in df_deals.iterrows():
            key = row['match_key']
            candidates = df_seibi[(df_seibi['match_key'] == key) & (~df_seibi['orig_index'].isin(seibi_matched_indices))]
            if len(candidates) > 0:
                matched_idx = candidates.index[0]
                seibi_matched_indices.add(df_seibi.loc[matched_idx, 'orig_index'])
            else:
                deals_unmatched.append(row)

        deals_unmatched_df = pd.DataFrame(deals_unmatched)
        seibi_unmatched_df = df_seibi[~df_seibi['orig_index'].isin(seibi_matched_indices)].copy()

        # --- 出力データ整形 ---
        out_deals = pd.DataFrame({
            '存在ファイル': 'deals.csvのみ',
            '車両登録番号': deals_unmatched_df['備考'] if not deals_unmatched_df.empty else [],
            '売上種別': deals_unmatched_df.apply(lambda r: '車検' if (pd.isna(r['品目']) and r['勘定科目']=='車検') else r['品目'], axis=1) if not deals_unmatched_df.empty else [],
            '金額（税抜）': deals_unmatched_df['amount_clean'] if not deals_unmatched_df.empty else [],
            '伝票番号/管理番号': deals_unmatched_df['管理番号'] if not deals_unmatched_df.empty else [],
            '取引先/請求先名': deals_unmatched_df['取引先'] if not deals_unmatched_df.empty else [],
            '勘定科目': deals_unmatched_df['勘定科目'] if not deals_unmatched_df.empty else [],
            '発生日/売上日': deals_unmatched_df['発生日'] if not deals_unmatched_df.empty else []
        })

        out_seibi = pd.DataFrame({
            '存在ファイル': '整備売上明細CSVのみ',
            '車両登録番号': seibi_unmatched_df['登録番号'],
            '売上種別': seibi_unmatched_df['種別'],
            '金額（税抜）': seibi_unmatched_df['amount_clean'],
            '伝票番号/管理番号': seibi_unmatched_df['伝票No'].astype(str),
            '取引先/請求先名': seibi_unmatched_df['請求先名'],
            '勘定科目': '-',
            '発生日/売上日': seibi_unmatched_df['売上日']
        })

        result_df = pd.concat([out_deals, out_seibi], ignore_index=True)

        # --- 結果表示 ---
        st.markdown("---")
        st.subheader("📊 照合結果サマリー")
        
        m1, m2, m3 = st.columns(3)
        m1.metric("不一致データ総数", f"{len(result_df)} 件")
        m2.metric("deals.csv のみ存在", f"{len(out_deals)} 件")
        m3.metric("整備売上明細CSV のみ存在", f"{len(out_seibi)} 件")

        if not result_df.empty:
            st.dataframe(result_df, use_container_width=True)

            csv_data = result_df.to_csv(index=False, encoding='utf-8-sig')
            st.download_button(
                label="📥 不一致データ一覧（CSV）をダウンロード",
                data=csv_data,
                file_name="不一致データ一覧.csv",
                mime="text/csv"
            )
        else:
            st.success("🎉 すべてのデータが一致しました！不一致データはありません。")

    except Exception as e:
        st.error(f"⚠️ 処理中にエラーが発生しました: {e}")
        st.info("ファイル形式や内容に異常がないかご確認ください。")
