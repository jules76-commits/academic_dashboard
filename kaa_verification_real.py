import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="KAA Qualification Verification", layout="wide")
st.title("📚 KAA Employee Academic Qualification Dashboard")

uploaded_file = st.file_uploader("📤 Upload the REAL Excel file with employee qualifications", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Load file with clean headers
        df = pd.read_excel(uploaded_file)
        df.columns = df.columns.str.strip().str.upper()

        # Show columns for confirmation
        st.write("📋 Columns detected:")
        for i, col in enumerate(df.columns):
            st.write(f"{i + 1}. '{col}'")

        # Add missing columns if necessary
        if 'VERIFIED' not in df.columns:
            df['VERIFIED'] = 'Pending'
        if 'REMARKS' not in df.columns:
            df['REMARKS'] = ''

        st.write("### 🧾 Employee Verification Table")

        for i in range(len(df)):
            st.markdown("---")
            st.markdown(f"### 👤 {df.at[i, 'FULL NAME']} — *{df.at[i, 'POSITION']}*")
            st.markdown(f"🆔 Personal Number: `{df.at[i, 'PERSONAL NUMBER']}`")

            with st.expander("📄 View Academic Qualifications", expanded=False):
                st.write({
                    "PhD": df.at[i, 'PHD'],
                    "Masters": df.at[i, 'MASTERS'],
                    "Undergraduate": df.at[i, 'UNDERGRADUATE'],
                    "Diploma": df.at[i, 'DIPLOMA'],
                    "Certificate": df.at[i, 'CERTIFICATE'],
                    "KCSE": df.at[i, 'KCSE']
                })

            df.at[i, 'VERIFIED'] = st.selectbox(
                "✅ Verification Status",
                ['Pending', 'Verified', 'Rejected'],
                index=['Pending', 'Verified', 'Rejected'].index(df.at[i, 'VERIFIED']),
                key=f"verify_{i}"
            )

            df.at[i, 'REMARKS'] = st.text_input(
                "📝 Add Remarks",
                value=df.at[i, 'REMARKS'],
                key=f"remarks_{i}"
            )

        st.write("### ✅ Final Preview Table")
        st.dataframe(df)

        def to_excel(dataframe):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                dataframe.to_excel(writer, index=False)
            return output.getvalue()

        st.download_button(
            label="📥 Download Verified Report",
            data=to_excel(df),
            file_name="KAA_Qualified_Employees_Verified.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
else:
    st.info("👈 Upload the KAA employee Excel file to begin.")
