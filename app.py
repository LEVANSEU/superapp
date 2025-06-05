import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
import re

st.set_page_config(layout="wide")
st.title("Excel გენერატორი")

# გაფართოებული ვიზუალი
st.markdown("""
    <style>
        .main { max-width: 95%; padding-left: 2rem; padding-right: 2rem; }
        .block-container { padding-top: 1rem; padding-bottom: 1rem; }
        button[kind="secondary"] { width: 100%; }
    </style>
""", unsafe_allow_html=True)

report_file = st.file_uploader("ატვირთე ანგარიშფაქტურების ფაილი (report.xlsx)", type=["xlsx"])
statement_file = st.file_uploader("ატვირთე საბანკო ამონაწერის ფაილი (statement.xlsx)", type=["xlsx"])

if report_file and statement_file:
    purchases_df = pd.read_excel(report_file, sheet_name='Grid')
    bank_df = pd.read_excel(statement_file)

    purchases_df['დასახელება'] = purchases_df['გამყიდველი'].astype(str).apply(lambda x: re.sub(r'^\(\d+\)\s*', '', x).strip())
    purchases_df['საიდენტიფიკაციო კოდი'] = purchases_df['გამყიდველი'].apply(lambda x: ''.join(re.findall(r'\d', str(x)))[:11])
    bank_df['P'] = bank_df.iloc[:, 15].astype(str).str.strip()
    bank_df['Amount'] = pd.to_numeric(bank_df.iloc[:, 3], errors='coerce').fillna(0)

    wb = Workbook()
    wb.remove(wb.active)

    ws1 = wb.create_sheet(title="ანგარიშფაქტურები კომპანიით")
    ws1.append(['დასახელება', 'საიდენტიფიკაციო კოდი', 'ანგარიშფაქტურის №', 'ანგარიშფაქტურის თანხა', 'ჩარიცხული თანხა'])

    company_summaries = []

    for company_id, group in purchases_df.groupby('საიდენტიფიკაციო კოდი'):
        company_name = group['დასახელება'].iloc[0]
        unique_invoices = group.groupby('სერია №')['ღირებულება დღგ და აქციზის ჩათვლით'].sum().reset_index()
        company_invoice_sum = unique_invoices['ღირებულება დღგ და აქციზის ჩათვლით'].sum()

        company_summary_row = ws1.max_row + 1
        payment_formula = f"=SUMIF(საბანკოამონაწერი!P:P, B{company_summary_row}, საბანკოამონაწერი!D:D)"
        ws1.append([company_name, company_id, '', company_invoice_sum, payment_formula])

        for _, row in unique_invoices.iterrows():
            ws1.append(['', '', row['სერია №'], row['ღირებულება დღგ და აქციზის ჩათვლით'], ''])

        company_summaries.append((company_name, company_id, company_invoice_sum))

    for sheet_title, content_df in [
        ("დეტალური მონაცემები", purchases_df),
        ("საბანკოამონაწერი", bank_df),
        ("ანგარიშფაქტურის დეტალები", purchases_df[['სერია №', 'საქონელი / მომსახურება', 'ზომის ერთეული', 'რაოდ.', 'ღირებულება დღგ და აქციზის ჩათვლით']].rename(columns={'სერია №': 'ინვოისის №'})),
        ("გადარიცხვები_უბმოლოდ", bank_df[~bank_df['P'].isin(purchases_df['საიდენტიფიკაციო კოდი'])]),
        ("განახლებული ამონაწერი", bank_df),
    ]:
        ws = wb.create_sheet(title=sheet_title)
        ws.append(content_df.columns.tolist())
        for row in content_df.itertuples(index=False):
            ws.append(row)

    ws7 = wb.create_sheet(title="კომპანიების_ჯამები")
    ws7.append(['დასახელება', 'საიდენტიფიკაციო კოდი', 'ანგარიშფაქტურების ჯამი', 'ჩარიცხული თანხა'])
    for idx, (company_name, company_id, invoice_sum) in enumerate(company_summaries, start=2):
        payment_formula = f"=SUMIF(საბანკოამონაწერი!P:P, B{idx}, საბანკოამონაწერი!D:D)"
        ws7.append([company_name, company_id, invoice_sum, payment_formula])

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    if 'selected_company' not in st.session_state:
        st.subheader("📋 კომპანიების ჩამონათვალი")

        for name, company_id, invoice_sum in company_summaries:
            col1, col2, col3, col4, col5 = st.columns([2, 2, 1.5, 1.5, 1.5])
            with col1:
                if st.button(f"{name}", key=f"name_{company_id}"):
                    st.session_state['selected_company'] = name
            with col2:
                if st.button(f"{company_id}", key=f"id_{company_id}"):
                    st.session_state['selected_company'] = name

            paid_sum = bank_df[bank_df["P"] == str(company_id)]["Amount"].sum()
            difference = invoice_sum - paid_sum

            with col3:
                st.write(f"{invoice_sum:,.2f}")
            with col4:
                st.write(f"{paid_sum:,.2f}")
            with col5:
                st.write(f"{difference:,.2f}")

    else:
        selected_name = st.session_state['selected_company']
        st.subheader(f"🔎 {selected_name} - ანგარიშფაქტურები")

        report_file.seek(0)
        df_full = pd.read_excel(report_file, sheet_name='Grid')
        df_full['დასახელება'] = df_full['გამყიდველი'].astype(str).apply(lambda x: re.sub(r'^\(\d+\)\s*', '', x).strip())
        matching_df = df_full[df_full['დასახელება'].str.contains(selected_name, na=False)].copy()
        matching_df['რაოდ.'] = matching_df['რაოდ.'].fillna(1)

        if not matching_df.empty:
            st.subheader("🧮 ინვოისის გადათვლა - ახალი ფასებით")

            if 'new_prices' not in st.session_state:
                st.session_state.new_prices = {}

            total = 0
            for i, row in matching_df.iterrows():
                col1, col2, col3, col4, col5 = st.columns([4, 2, 1, 2, 2])
                with col1:
                    st.markdown(row['საქონელი / მომსახურება'])
                with col2:
                    st.markdown(row['ზომის ერთეული'])
                with col3:
                    qty = row['რაოდ.'] or 1
                    st.markdown(str(qty))
                with col4:
                    key = f"new_price_{i}"
                    new_price = st.number_input("ახალი ფასი", value=st.session_state.new_prices.get(key, 0.0), key=key, format="%.2f")
                    st.session_state.new_prices[key] = new_price
                with col5:
                    use_price = new_price if new_price > 0 else row['ღირებულება დღგ და აქციზის ჩათვლით']
                    item_total = qty * use_price
                    st.markdown(f"**{item_total:.2f} ₾**")
                    total += item_total

            st.markdown("---")
            st.subheader(f"📊 ჯამი: **{total:.2f} ₾**")

            st.subheader("🔍 მოძებნე გუგლში მასალა ან მომსახურება")
            col1, col2 = st.columns([3, 1])
            with col1:
                search_term = st.text_input("ჩაწერე სახელი ან სიტყვა:")
            with col2:
                if st.button("ძებნა"):
                    if search_term.strip():
                        search_url = f"https://www.google.com/search?q={search_term.replace(' ', '+')}"
                        st.markdown(f"[🌐 გადადი გუგლზე]({search_url})", unsafe_allow_html=True)
                    else:
                        st.warning("გთხოვ ჩაწერე ტექსტი ძებნამდე.")

            # Excel ჩამოტვირთვა
            company_output = io.BytesIO()
            company_wb = Workbook()
            ws = company_wb.active
            ws.title = selected_name[:31]
            ws.append(matching_df.columns.tolist())
            for row in matching_df.itertuples(index=False):
                ws.append(row)
            company_wb.save(company_output)
            company_output.seek(0)

            st.download_button(
                label=f"⬇️ ჩამოტვირთე {selected_name} ინვოისების Excel",
                data=company_output,
                file_name=f"{selected_name}_ინვოისები.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        else:
            st.warning("📭 ჩანაწერი ვერ მოიძებნა ამ კომპანიისთვის.")

        if st.button("⬅️ დაბრუნება სრულ სიაზე"):
            del st.session_state['selected_company']

    st.success("✅ ფაილი მზადაა! ჩამოტვირთე აქედან:")
    st.download_button(
        label="⬇️ ჩამოტვირთე Excel ფაილი",
        data=output,
        file_name="საბოლოო_ფაილი.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
