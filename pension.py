import streamlit as st
import pandas as pd
import qrcode
import io

# ===============================
#   QR CODE 產生函式（可重用）
# ===============================
def generate_qr(url: str):
    """產生 QR Code 並以 bytes 回傳"""
    qr = qrcode.make(url)
    buf = io.BytesIO()
    qr.save(buf, format="PNG")
    return buf.getvalue()   # 回傳 byte image

# ===============================
#   Streamlit UI
# ===============================
st.set_page_config(page_title="退休金決策系統 Retirement App", layout="wide")

st.title("退休金計畫互動系統 · Retirement Income App")

st.write("本系統示範如何在 Streamlit 中嵌入 QR Code，用於 Windows App 或手機版連結。")

# === 你的 Cloud App 或外部連結（請改這一行）===
app_url = "https://2025-cross-disciplinary-creative-programming-competition-6lnde.streamlit.app/"

# 產生 QR Code
qr_img = generate_qr(app_url)

# 顯示 QR Code
st.subheader("📱 手機掃描 QR Code 開啟 App")
st.image(qr_img, width=240, caption=app_url)

# ===============================
def calc_pension_model(
    b_prepaid,
    pbo_input,
    b_pa,
    b_prior,
    b_gl,
    b_trans,
    rate,
    ret_rate,
    amortY_gl,
    amortY_prior,
    amortY_trans,
    manual_actuarial_amort,
    adj_prior_gl,
    adj_prior,
    svc_cost,
    actual_ret_input,
    pbo_gl,
    adj_shinks,
    contrib_base_input,
    contrib_pension,
):
    """
    Pension model (X9)
    - 若發生方法變動（前期服務成本 / 精算損益調整）
      → 期初預付(應付)退休金須先重編
    - 其餘年度維持原期初數
    """

    # === 期初餘額（符號轉換）===
    b_pbo = -pbo_input                  # PBO 為負債
    actual_ret = -actual_ret_input      # 實際報酬為減項
    contrib_base = -contrib_base_input  # 公司提撥為現金流出

    # === 中間計算 ===
    expected_ret = -((b_pa * ret_rate) + actual_ret)
    amort_prior = (b_prior + adj_prior) / amortY_prior if amortY_prior != 0 else 0

    # 精算損益攤銷：有年限→模型；沒年限→人工
    if amortY_gl and amortY_gl > 0:
        amort_actuarial_gl = (-expected_ret + pbo_gl - amort_prior) / amortY_gl
    else:
        amort_actuarial_gl = manual_actuarial_amort

    # === 調整後 PBO（方法變動、縮減清償都影響）===
    adj_b_pbo = b_pbo - adj_prior - adj_shinks - adj_prior_gl
    interest_cost = -adj_b_pbo * rate
    amort_trans = b_trans / amortY_trans if amortY_trans != 0 else 0

    # === 本期退休金費用 ===
    pension_expense = (
        svc_cost
        + interest_cost
        + actual_ret
        + expected_ret
        + amort_actuarial_gl
        + amort_prior
        + amort_trans
    )

    # ================================
    # ★ 關鍵修正：期初退休金是否重編
    # ================================
    if adj_prior != 0 or adj_prior_gl != 0:
        # 期初五大構成項「重編後」再加總
        adj_b_pa    = b_pa
        adj_b_prior = b_prior + adj_prior
        adj_b_gl    = b_gl + adj_prior_gl
        adj_b_trans = b_trans

        b_prepaid_effective = (
            adj_b_pbo
            + adj_b_pa
            + adj_b_prior
            + adj_b_gl
            + adj_b_trans
        )
    else:
        b_prepaid_effective = b_prepaid

    # === 期末餘額 ===
    end_pbo = adj_b_pbo - svc_cost - interest_cost - pbo_gl + contrib_pension
    end_pa = b_pa - actual_ret - contrib_base - contrib_pension
    end_prior = b_prior - amort_prior + adj_prior
    end_gl = b_gl - amort_actuarial_gl + pbo_gl + adj_prior_gl - expected_ret
    end_trans = b_trans - amort_trans
    end_prepaid = b_prepaid_effective - (pension_expense + contrib_base)

    # === 分錄（簡化 Dr / Cr）===
    cash = -contrib_base  # 現金流出（正數）

    journal_rows = []
    journal_rows.append({"Account": "退休金費用", "Debit": pension_expense, "Credit": 0})

    if cash != 0:
        journal_rows.append({"Account": "現金", "Debit": 0, "Credit": cash})

    if end_prepaid >= 0:
        diff = end_prepaid - b_prepaid_effective
        if diff != 0:
            journal_rows.append({"Account": "預付退休金", "Debit": diff, "Credit": 0})
    else:
        diff = -(end_prepaid - b_prepaid_effective)
        if diff != 0:
            journal_rows.append({"Account": "應付退休金", "Debit": 0, "Credit": diff})

    # === 成本拆解 ===
    expense_detail = [
        ("當期服務成本", svc_cost),
        ("利息費用", interest_cost),
        ("實際報酬（減項）", actual_ret),
        ("計畫資產損益（減項）", expected_ret),
        ("精算損益攤銷", amort_actuarial_gl),
        ("前期服務成本攤銷", amort_prior),
        ("過渡性淨負債攤銷", amort_trans),
    ]

    # === 期末五大構成項 ===
    ending_balances = {
        "預付/(應付)退休金": end_prepaid,
        "確定給付義務現值": end_pbo,
        "計畫資產": end_pa,
        "未認列前期服務成本": end_prior,
        "未認列精算損益": end_gl,
        "未認列過渡性淨負債": end_trans,
    }

    return expense_detail, pension_expense, ending_balances, journal_rows

# ================= Streamlit 介面 =================

st.set_page_config(page_title="Pension Plan – X9 模擬", layout="wide")

st.title("退休金計畫模擬器（Pension Plan Simulator）")
st.caption("對齊 1122 AIS-game.xlsx / Pension 工作表：X9 年退休金費用、期末餘額與正式分錄")

st.markdown("---")

with st.form("pension_form"):
    c1, c2, c3 = st.columns(3)

    # ---- ① X9 期初餘額 ----
    with c1:
        st.subheader("① X9 年初餘額")
        b_prepaid = st.number_input("預付(+)/應付(-)退休金", value=-170000.0, step=1000.0)
        pbo_input = st.number_input("確定給付義務（負向，輸入正數）", value=1500000.0, step=10000.0)
        b_pa = st.number_input("退休金計畫資產", value=900000.0, step=10000.0)
        b_prior = st.number_input("未認列前期服務成本", value=0.0, step=10000.0)
        b_gl = st.number_input("未認列精算損益（利益為+）", value=180000.0, step=10000.0)
        b_trans = st.number_input("未認列過渡性淨負債", value=250000.0, step=10000.0)

    # ---- ② 參數與攤銷年限 ----
    with c2:
        st.subheader("② 參數與攤銷年限")
        rate = st.number_input("殖利率 (%)", value=10.0, step=0.5) / 100
        ret_rate = st.number_input("退休基金報酬率 (%)", value=10.0, step=0.5) / 100
        amortY_gl = st.number_input("精算損益攤銷年限（=0 表示不自動攤銷，改用下方人工金額）", value=20.0, step=1.0)
        amortY_prior = st.number_input("前期服務成本攤銷年限（=0 表示不攤銷）", value=20.0, step=1.0)
        amortY_trans = st.number_input("過渡性淨負債攤銷年限（=0 表示不攤銷）", value=4.0, step=1.0)
        contrib_base_input = st.number_input("本期提撥數（提撥計畫資產）", value=350000.0, step=10000.0)
        contrib_pension = st.number_input("本期支付退休金（付給退休員工）", value=300000.0, step=10000.0)

    # ---- ③ X9 本期變動 ----
    with c3:
        st.subheader("③ X9 年本期交易")
        adj_prior_gl = st.number_input("方法變動調整未認列精算損益(利益為+)", value=0.0, step=10000.0)
        adj_prior = st.number_input("方法變動調整前期服務成本", value=200000.0, step=10000.0)
        svc_cost = st.number_input("當期服務成本", value=200000.0, step=10000.0)
        actual_ret_input = st.number_input("實際報酬（輸入正數，程式自動當成減項）", value=70000.0, step=10000.0)
        manual_actuarial_amort = st.number_input(
            "精算損益攤銷（人工輸入；若上面年限=0 或留空時使用）",
            value=0.0,
            step=10000.0
        )
        pbo_gl = st.number_input("確定給付義務現值損益（利益為+）", value=0.0, step=10000.0)
        adj_shinks = st.number_input("縮減或清償損益（利益為+）", value=0.0, step=10000.0)

    submitted = st.form_submit_button("開始計算（Run Simulation）")

if submitted:
    try:
        expense_detail, pension_expense, ending_balances, journal_rows = calc_pension_model(
            b_prepaid=b_prepaid,
            pbo_input=pbo_input,
            b_pa=b_pa,
            b_prior=b_prior,
            b_gl=b_gl,
            b_trans=b_trans,
            rate=rate,
            ret_rate=ret_rate,
            amortY_gl=amortY_gl,
            amortY_prior=amortY_prior,
            amortY_trans=amortY_trans,
            manual_actuarial_amort=manual_actuarial_amort,  # ★ 關鍵：把它傳進去
            adj_prior_gl=adj_prior_gl,
            adj_prior=adj_prior,
            svc_cost=svc_cost,
            actual_ret_input=actual_ret_input,
            pbo_gl=pbo_gl,
            adj_shinks=adj_shinks,
            contrib_base_input=contrib_base_input,
            contrib_pension=contrib_pension,
        )
    except Exception as e:
        st.error(f"計算過程發生錯誤：{e}")
    else:
        resL, resR = st.columns(2)

        with resL:
            st.subheader("④ 本期退休金費用計算")
            df_exp = pd.DataFrame(
                [{"項目": name, "金額": amt} for name, amt in expense_detail]
                + [{"項目": "退休金費用合計", "金額": pension_expense}]
            )
            st.table(df_exp.style.format({"金額": "{:,.0f}"}))

        with resR:
            st.subheader("⑤ 期末餘額與正式分錄")

            df_bal = pd.DataFrame(
                [{"項目": k, "期末餘額": v} for k, v in ending_balances.items()]
            )
            st.markdown("**期末餘額（Ending Balances）**")
            st.table(df_bal.style.format({"期末餘額": "{:,.0f}"}))

            st.markdown("**X9 年正式分錄（Journal Entry）**")
            df_je = pd.DataFrame(journal_rows)
            st.table(df_je.style.format({"Debit": "{:,.0f}", "Credit": "{:,.0f}"}))

        st.success("計算完成，已對齊 Excel Pension 範例的邏輯。")
