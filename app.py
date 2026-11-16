import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Network Materials Supply Planning",
    page_icon="⚡",
    layout="wide",
)

st.title("📦 نظام تخطيط احتياجات الشبكات وإدارة المخزون والموردين")
st.markdown(
    """
> **نظام ذكي متكامل لسلاسل الإمداد وتخطيط المخزون (تصميم خاص لمقابلات إدارة الإمداد والشبكات).**
"""
)

excel_source = "SupplyChain_Interview_Project.xlsx"


@st.cache_data
def load_data(file):
  dash = pd.read_excel(file, sheet_name="Dashboard")
  items = pd.read_excel(file, sheet_name="Items", skiprows=2)
  analysis = pd.read_excel(file, sheet_name="Analysis", skiprows=2)
  suppliers = pd.read_excel(file, sheet_name="Suppliers", skiprows=2)
  return dash, items, analysis, suppliers


try:
  dash_df, items_df, analysis_df, suppliers_df = load_data(excel_source)

  st.sidebar.header("🧭 لوحة التحكم الرئيسية")
  menu = st.sidebar.selectbox(
      "اختر القسم:",
      [
          "لوحة المؤشرات الرئيسية (Dashboard)",
          "تحليل المخزون وإعادة الطلب (Analysis)",
          "المحاكي التفاعلي اللحظي (Interactive Simulator)",
          "تقييم الموردين (Suppliers Scorecard)",
      ],
  )

  if menu == "لوحة المؤشرات الرئيسية (Dashboard)":
    st.subheader("📊 المؤشرات الرئيسية لأداء المخزون والصرف")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
      st.metric(label="إجمالي الصرف السنوي", value="250,043,180 EGP")
    with col2:
      st.metric(label="قيمة المخزون الحالي", value="55,443,560 EGP")
    with col3:
      st.metric(label="معدل دوران المخزون", value="4.51 x")
    with col4:
      st.metric(label="أصناف تتطلب إعادة طلب", value="5 أصناف", delta="حر")

    st.markdown("---")
    st.subheader("📋 ملخص حالة الأصناف في النظام")
    st.dataframe(
        analysis_df[
            [
                "Item ID",
                "Item",
                "ABC-XYZ",
                "Status",
                "Suggested order qty",
                "Months of cover",
            ]
        ],
        use_container_width=True,
    )

  elif menu == "تحليل المخزون وإعادة الطلب (Analysis)":
    st.subheader("⚙️ جدول تحليل المخزون، نقطة إعادة الطلب (ROP) و EOQ")
    st.dataframe(analysis_df, use_container_width=True)

  elif menu == "المحاكي التفاعلي اللحظي (Interactive Simulator)":
    st.subheader("🎛️ محاكي تخطيط الاحتياجات ولحظي التغيير")
    col_s1, col_s2 = st.columns(2)
    with col_s1:
      avg_demand = st.number_input(
          "متوسط الطلب الشهري (Units)", min_value=1, value=500
      )
      lead_time_days = st.slider(
          "فترة التوريد (Lead Time in Days)", min_value=5, max_value=120, value=30
      )
    with col_s2:
      unit_cost = st.number_input("تكلفة الوحدة (EGP)", min_value=1.0, value=1500.0)

    lead_time_months = lead_time_days / 30.0
    safety_stock = int(avg_demand * 0.15 * 1.65)
    reorder_point = int((avg_demand * lead_time_months) + safety_stock)
    eoq = int((2 * avg_demand * 12 * 1500 / (unit_cost * 0.2)) ** 0.5)

    st.markdown("---")
    res_col1, res_col2, res_col3 = st.columns(3)
    with res_col1:
      st.metric(label="مخزون الأمان المقترح", value=f"{safety_stock} وحدة")
    with res_col2:
      st.metric(label="نقطة إعادة الطلب (ROP)", value=f"{reorder_point} وحدة")
    with res_col3:
      st.metric(label="كمية الطلب الاقتصادية (EOQ)", value=f"{eoq} وحدة")

  elif menu == "تقييم الموردين (Suppliers Scorecard)":
    st.subheader("🤝 لوحة قياس أداء الموردين (Supplier Scorecard)")
    st.dataframe(suppliers_df, use_container_width=True)

except Exception as e:
  st.error(f"حدث خطأ أثناء تحميل البيانات: {e}")