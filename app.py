import streamlit as st
import pandas as pd
import io

def calculate_sumifs(file, file_ext):
    try:
        # ==== Đọc file Excel ====
        if file_ext == ".xlsb":
            df = pd.read_excel(file, engine="pyxlsb")
        else:  # .xlsx hoặc .xls
            df = pd.read_excel(file, sheet_name=0)

        df.columns = df.columns.str.strip()

        # ==== B1. Cột Past due → Total_Demand ====
        main_cols = df.loc[:, 'Past due':'Total_Demand'].columns.tolist()

        # ==== B2. Group Firm+Forecast ====
        filtered = df[df['Type'].isin(['Firm', 'Forecast'])]
        grouped = (
            filtered.groupby(['Part_No', 'Vendor_Code'])[main_cols]
            .sum()
            .reset_index()
        )
        grouped['Type'] = "Firm+Forecast"

        # ==== B3. Store_Qty & IQC_QTY theo site ====
        firm_rows = df[df['Type'] == 'Firm'].copy()

        s1 = (
            firm_rows[firm_rows['Site'] == 'TH3-SHTP']
            .groupby(['Part_No','Vendor_Code'])[['Store_Qty','IQC_QTY']]
            .sum()
            .reset_index()
            .rename(columns={'Store_Qty':'Store_TH3','IQC_QTY':'IQC_TH3'})
        )

        s2 = (
            firm_rows[firm_rows['Site'] == 'TD3-DDK']
            .groupby(['Part_No','Vendor_Code'])[['Store_Qty','IQC_QTY']]
            .sum()
            .reset_index()
            .rename(columns={'Store_Qty':'Store_TD3','IQC_QTY':'IQC_TD3'})
        )

        store_qty = pd.merge(s1, s2, on=['Part_No','Vendor_Code'], how='outer').fillna(0)
        store_qty['Store_Qty'] = store_qty['Store_TH3'] + store_qty['Store_TD3']
        store_qty['IQC_QTY'] = store_qty['IQC_TH3'] + store_qty['IQC_TD3']
        store_qty = store_qty[['Part_No','Vendor_Code','Store_Qty','IQC_QTY']]

        # ==== B4. Metadata ====
        meta_cols = [c for c in ['Part_No','Vendor_Code','Buyer','Planner','Vendor','Org','Site'] if c in df.columns]
        meta_info = df[meta_cols].drop_duplicates(subset=['Part_No','Vendor_Code'])

        # ==== B5. Merge ====
        result = pd.merge(grouped, store_qty, on=['Part_No','Vendor_Code'], how='left')
        result = pd.merge(result, meta_info, on=['Part_No','Vendor_Code'], how='left')

        # ==== B6. Reorder columns ====
        cols = result.columns.tolist()
        new_order = ['Part_No','Vendor_Code','Type']

        if 'Store_Qty' in cols: 
            new_order.append('Store_Qty')
        if 'IQC_QTY' in cols: 
            new_order.append('IQC_QTY')

        new_order += [c for c in main_cols if c in cols and c != 'Total_Demand']
        new_order += [c for c in cols if c not in new_order]

        result = result[new_order]

        # ==== B6.1 Xóa cột Total_Demand và Site ====
        drop_cols = [c for c in ['Total_Demand','Site'] if c in result.columns]
        result = result.drop(columns=drop_cols)

        return result

    except Exception as e:
        st.error(f"Lỗi khi xử lý: {str(e)}")
        return None


# ==== Streamlit UI ====
st.set_page_config(page_title="📊 Firm+Forecast Tins", layout="centered")

st.title("📊 Firm + Forecast Data Tool")

uploaded_file = st.file_uploader("Chọn file Excel", type=["xlsx", "xls", "xlsb"])

if uploaded_file is not None:
    file_ext = "." + uploaded_file.name.split(".")[-1].lower()

    if st.button("Xử lý dữ liệu"):
        result = calculate_sumifs(uploaded_file, file_ext)
        if result is not None:
            st.success("✅ Đã xử lý dữ liệu thành công!")
            st.dataframe(result.head(20))

            # Xuất file Excel kết quả ra bộ nhớ tạm
            out_name = uploaded_file.name.replace(file_ext,"_FirmForecast_Sum.xlsx")
            towrite = io.BytesIO()
            result.to_excel(towrite, index=False, engine='openpyxl')
            towrite.seek(0)

            st.download_button(
                label="📥 Tải kết quả Excel",
                data=towrite,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
