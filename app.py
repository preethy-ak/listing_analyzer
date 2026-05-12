import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(layout="wide")
st.title("📊 Marketplace Listing Audit Dashboard")

TODAY = datetime.today()

# -------------------------------
# UPLOAD
# -------------------------------
st.sidebar.header("Upload Files")

zecom = st.sidebar.file_uploader("Zecom Tracker", type=["xlsx"])
lazada = st.sidebar.file_uploader("Lazada File", type=["xlsx"])
shopee = st.sidebar.file_uploader("Shopee File", type=["xlsx"])
zalora = st.sidebar.file_uploader("Zalora File", type=["xlsx"])
inventory = st.sidebar.file_uploader("Inventory File", type=["csv"])

# -------------------------------
# NORMALIZE FUNCTION
# -------------------------------
def norm(x):
    try:
        return str(int(float(x))).strip()
    except:
        return str(x).strip()

# -------------------------------
# AGGREGATION FUNCTION
# -------------------------------
def process_marketplace(df, sku_col, status_col):
    df["SKU"] = df[sku_col].apply(norm)
    df["status"] = df[status_col].astype(str).str.lower().str.strip()

    agg = df.groupby("SKU").agg({
        "status": lambda x: "active" if "active" in list(x) else "inactive"
    }).reset_index()

    return agg.set_index("SKU")["status"].to_dict()

# -------------------------------
# MAIN LOGIC
# -------------------------------
if zecom and inventory:

    df_content = pd.read_excel(zecom, sheet_name="Content file")
    df_content["EAN"] = df_content["EAN"].apply(norm)

    df_inv = pd.read_csv(inventory)
    df_inv["EAN"] = df_inv["EAN"].apply(norm)

    inv_dict = df_inv.set_index("EAN")["Avail_Qty"].to_dict()

    # -------------------------------
    # LOAD MARKETPLACES
    # -------------------------------
    laz_dict, shp_dict, zal_dict = {}, {}, {}

    if lazada:
        df_laz = pd.read_excel(lazada)
        laz_dict = process_marketplace(df_laz, "SellerSKU", "Status")

    if shopee:
        df_shp = pd.read_excel(shopee)
        shp_dict = process_marketplace(df_shp, "SellerSku", "Status")

    if zalora:
        df_zal = pd.read_excel(zalora)
        zal_dict = process_marketplace(df_zal, "SellerSku", "Status")

    # -------------------------------
    # BUILD FINAL DATA
    # -------------------------------
    final = []

    for _, row in df_content.iterrows():
        ean = row["EAN"]
        stock = inv_dict.get(ean, 0)

        laz = laz_dict.get(ean, "not_listed")
        shp = shp_dict.get(ean, "not_listed")
        zal = zal_dict.get(ean, "not_listed")

        statuses = [laz, shp, zal]

        launch_date = row.get("Launch Date")
        if pd.notna(launch_date):
            launch_date = pd.to_datetime(launch_date)
        else:
            launch_date = None

        # -------------------------------
        # BUSINESS LOGIC
        # -------------------------------
        if launch_date and launch_date > TODAY:
            final_status = "Not Live Yet"
            comment = "Future launch"

        elif stock == 0:
            if all(s == "not_listed" for s in statuses):
                final_status = "OK"
                comment = "No stock, not listed (correct)"
            else:
                final_status = "Inactive Required"
                comment = "Stock 0 but listing active"

        else:
            if all(s == "not_listed" for s in statuses):
                final_status = "Action Needed"
                comment = "Stock available but not listed"
            elif any(s == "active" for s in statuses):
                final_status = "Active"
                comment = "Live on marketplace"
            else:
                final_status = "Inactive"
                comment = "Listed but inactive"

        final.append({
            "EAN": ean,
            "Product Name": row.get("Product Name", ""),
            "Stock": stock,
            "Lazada": laz,
            "Shopee": shp,
            "Zalora": zal,
            "Final Status": final_status,
            "Comment": comment
        })

    df_final = pd.DataFrame(final)

    # -------------------------------
    # DASHBOARD
    # -------------------------------
    st.subheader("📊 Summary")

    col1, col2, col3, col4 = st.columns(4)

    col1.metric("Total SKUs", len(df_final))
    col2.metric("Active", (df_final["Final Status"] == "Active").sum())
    col3.metric("Action Needed", (df_final["Final Status"] == "Action Needed").sum())
    col4.metric("Issues", df_final["Final Status"].isin(["Inactive","Inactive Required"]).sum())

    # -------------------------------
    # BREAKDOWN
    # -------------------------------
    st.subheader("📌 Status Breakdown")
    st.bar_chart(df_final["Final Status"].value_counts())

    # -------------------------------
    # FILTER
    # -------------------------------
    status_filter = st.multiselect(
        "Filter by Status",
        df_final["Final Status"].unique(),
        default=df_final["Final Status"].unique()
    )

    st.dataframe(df_final[df_final["Final Status"].isin(status_filter)])

    # -------------------------------
    # DOWNLOAD
    # -------------------------------
    st.download_button(
        "⬇️ Download Report",
        df_final.to_csv(index=False),
        file_name="marketplace_audit.csv"
    )

else:
    st.info("Upload Zecom + Inventory (mandatory). Add marketplace files optionally.")
