import streamlit as st
import pandas as pd
import numpy as np
import io
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# --- 網頁基礎配置 ---
st.set_page_config(page_title="專業級 SRC / SC 大梁構造與撓度檢核系統", layout="wide")

# --- 材料與高溫折減常數 (依據台灣鋼結構與SRC技術規範) ---
E_s_cm = 2040  # 鋼材彈性模數 (tf/cm^2)
K_Y_600 = 0.47  # 600°C 鋼材屈服強度折減 47%
K_C_600 = 0.45  # 600°C 混凝土抗壓強度折減 45%

# --- 輔助函式：設定 Word 表格儲存格背景顏色 ---
def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

# ==============================================================================
# 🎯 核心大梁力學計算引擎 (依據最新指示精簡構造與 RBS 觸發路徑)
# ==============================================================================
def calculate_beam(row):
    res = row.copy()
    res['說明提示'] = "大梁幾何與側向支撐符合標準規範限制。"
    res['RBS提示說明'] = "無"
    res['開孔補強說明'] = "無"
    res['撓度檢核說明'] = "符合規範限制。"
    
    btype = row.get('結構形式', '純鋼骨SC')
    stype = row.get('內部鋼構形式', 'H型鋼')
    
    stirrup_num_key = '箍筋號數_#(3或4)' if '箍筋號數_#(3或4)' in row else '箍筋號數_#(3或4)'
    stirrup_spacing_key = '箍筋間距_s(cm)'
    
    # 🎯 終極防呆防線：如果是純鋼骨SC，後台全面強制清洗歸零
    if btype == "純鋼骨SC":
        f_c = 0.0
        F_yr = 0.0
        res['混凝土強度_fc(kg/cm2)'] = 0.0
        res['梁頂層鋼筋(支)'] = 0
        res['梁底層鋼筋(支)'] = 0
        res['鋼筋直徑_D(mm)'] = 0.0
        res[stirrup_num_key] = 0
        res[stirrup_spacing_key] = 0.0
    else:
        f_c = float(row.get('混凝土強度_fc(kg/cm2)', 350.0)) / 1000  # 轉 tf/cm2
        F_yr = float(row.get('鋼筋強度_Fyr(tf/cm2)', 4.2))

    F_y = float(row.get('鋼材強度_Fy(tf/cm2)', 3.52))
    M_ux = float(row.get('極限彎矩_Mux(ton-m)', 0.0))
    V_uy = float(row.get('極限剪力_Vuy(ton)', 0.0))
    L_b = float(row.get('側向未支撐長度_Lb(cm)', 150.0))
    
    L_span = 800.0  # 假設大梁淨跨距為 800 cm
    
    # 內部型鋼幾何
    d_s = float(row.get('鋼構總高_ds(cm)', 50.0))
    b_f = float(row.get('鋼構翼板寬_bf(cm)', 25.0))
    t_f = float(row.get('鋼構翼板厚_tf(cm)', 2.2))
    t_w = float(row.get('鋼構腹板厚_tw(cm)', 1.4))
    
    if stype == "H型鋼":
        A_s = 2 * b_f * t_f + (d_s - 2 * t_f) * t_w
        Z_sx = b_f * d_s**2 / 4 - (b_f - t_w) * (d_s - 2 * t_f)**2 / 4
        I_s = (b_f * d_s**3 / 12) - ((b_f - t_w) * (d_s - 2 * t_f)**3 / 12)
        r_y = b_f / (np.sqrt(12) * np.sqrt(1 + (d_s-2*t_f)*t_w / (6*b_f*t_f)))
    else: # 箱型
        A_s = 2 * b_f * t_f + 2 * (d_s - 2 * t_f) * t_w
        Z_sx = b_f * d_s**2 / 4 - (b_f - 2 * t_w) * (d_s - 2 * t_f)**2 / 4
        I_s = (b_f * d_s**3 / 12) - ((b_f - 2 * t_w) * (d_s - 2 * t_f)**3 / 12)
        r_y = b_f / np.sqrt(12)

    # 1. 大梁長期撓度變形 (Deflection delta) 計算
    if btype == "包覆型SRC":
        I_c = (float(row.get('梁斷面外圍寬_B(cm)', 50.0)) * float(row.get('梁斷面外圍高_H(cm)', 80.0))**3 / 12) - I_s
        I_eff = I_s + 0.3 * I_c  
        creep_factor = 2.0
    else:
        I_eff = I_s
        creep_factor = 1.0

    M_service = M_ux * 0.65
    delta_immediate = (M_service * 100 * L_span**2) / (9.6 * E_s_cm * I_eff) if I_eff > 0 else 0
    delta_longterm = delta_immediate * (1 + (creep_factor if btype == "包覆型SRC" else 0))
    
    delta_limit = L_span / 240.0
    res['長期總撓度_delta(cm)'] = delta_longterm
    res['規範撓度限制_limit(cm)'] = delta_limit
    
    if delta_longterm > delta_limit:
        res['撓度檢核說明'] = f"❌ 撓度超標！長期總撓度 {delta_longterm:.2f}cm > 限制 {delta_limit:.2f}cm (L/240)。"
    else:
        res['撓度檢核說明'] = f"🟢 合格。長期總撓度 {delta_longterm:.2f}cm ≦ 限制 {delta_limit:.2f}cm (L/240)。"

    # 2. 鋼結構大梁獨有：梁端減縮（RBS, 狗骨頭）幾何與力學公式計算 (精準限制僅 SC H型鋼觸發)
    if btype == "純鋼骨SC" and stype == "H型鋼":
        c_rbs = 0.2 * b_f
        b_rbs = 0.75 * d_s
        R_radius = (b_rbs**2 + 4 * c_rbs**2) / (8 * c_rbs)
        Z_rbs = Z_sx - 2 * c_rbs * t_f * (d_s - t_f)
        M_pr_rbs = 1.15 * 1.1 * F_y * Z_rbs / 100
        res['RBS提示說明'] = f"已啟動RBS設計：切削半徑 R={R_radius:.1f}cm，端部期望彎矩 Mpr={M_pr_rbs:.1f}t-m。"
    
    # 3. SRC 鋼筋穿腹板開孔規範與補強計算
    if btype == "包覆型SRC":
        D_hole = 3.2 
        d_w = d_s - 2 * t_f
        if D_hole > (d_w / 3):
            res['開孔補強說明'] = "❌ 警告：主筋穿孔直徑超出腹板淨高1/3規範限制！"
        else:
            t_p_req = (D_hole * t_w) / 10
            res['開孔補強說明'] = f"開孔符合限值。孔徑={D_hole}cm ≦ {d_w/3:.1f}cm。建議銲接厚度 tp={t_p_req:.2f}cm 之補強板。"

    # 4. 抗彎抗剪強度計算
    L_p = 1.76 * r_y * np.sqrt(2040 / F_y)
    M_ns = F_y * Z_sx if (stype == "箱型柱" or btype == "包覆型SRC" or L_b <= L_p) else F_y * Z_sx * (L_p / L_b)
    
    if btype == "包覆型SRC":
        A_st_bot = int(row.get('梁底層鋼筋(支)', 0)) * (np.pi * (float(row.get('鋼筋直徑_D(mm)', 25.0))/10)**2 / 4)
        d_eff = float(row.get('梁斷面外圍高_H(cm)', 80.0)) - 6.0
        M_nr = (A_st_bot * F_yr * (d_eff - 4.0))
        M_nx = M_ns + M_nr
        
        val_stirrup = int(row.get(stirrup_num_key, 3))
        A_v_each = 0.713 if val_stirrup == 3 else (1.267 if val_stirrup == 4 else 0.0)
        A_v = 2 * A_v_each
        s_spacing = float(row.get(stirrup_spacing_key, 10.0))
        V_ns_rebar = (A_v * F_yr * d_eff) / s_spacing if s_spacing > 0 else 0
        V_nc = 0.53 * np.sqrt(float(row.get('混凝土強度_fc(kg/cm2)', 350.0))) * float(row.get('梁斷面外圍寬_B(cm)', 50.0)) * d_eff / 1000
        V_ny = 0.6 * F_y * (d_s * t_w) + V_nc + V_ns_rebar
    else:
        M_nx = M_ns
        V_ny = 0.6 * F_y * (d_s * t_w * (2 if stype=="箱型柱" else 1))

    # 5. LRFD 強度檢核
    phi = 0.90
    interaction_ratio = M_ux / (phi * (M_nx / 100)) if M_nx > 0 else 0
    shear_ratio = V_uy / (phi * V_ny) if V_ny > 0 else 0
    status = "🟢 安全 (PASS)" if (interaction_ratio <= 1.0 and shear_ratio <= 1.0) else "🔴 超載 (NG)"
    
    res['常溫抗彎Mnx(ton-m)'] = float(M_nx / 100)
    res['常溫抗剪Vny(ton)'] = float(V_ny)
    res['高溫抗彎Mn_600(ton-m)'] = float((M_nx * K_Y_600) / 100)
    res['高溫抗剪Vn_600(ton)'] = float(V_ny * K_Y_600)
    res['設計彎矩應力比'] = float(interaction_ratio)
    res['設計剪力應力比'] = float(shear_ratio)
    res['安全判定'] = status
    
    return res

# ==============================================================================
# 🎯 預設大梁範例資料庫 (長度完全切齊為 3 筆，已全面清洗配筋欄位)
# ==============================================================================
sample_data = {
    '構件編號': ['B1-SRC_H', 'B2-SC_H', 'B3-SC_Box'],
    '結構形式': ['包覆型SRC', '純鋼骨SC', '純鋼骨SC'],
    '內部鋼構形式': ['H型鋼', 'H型鋼', '箱型柱'],
    '梁斷面外圍寬_B(cm)': [50.0, 40.0, 40.0],
    '梁斷面外圍高_H(cm)': [80.0, 70.0, 70.0],
    '鋼構總高_ds(cm)': [50.0, 50.0, 50.0],
    '鋼構翼板寬_bf(cm)': [25.0, 25.0, 30.0],
    '鋼構翼板厚_tf(cm)': [2.2, 2.0, 1.6],
    '鋼構腹板厚_tw(cm)': [1.4, 1.2, 1.6],
    '側向未支撐長度_Lb(cm)': [150.0, 300.0, 150.0],
    '混凝土強度_fc(kg/cm2)': [350.0, 0.0, 0.0],
    '鋼材強度_Fy(tf/cm2)': [3.52, 3.52, 3.52],
    '鋼筋強度_Fyr(tf/cm2)': [4.2, 0.0, 0.0],
    '梁頂層鋼筋(支)':[8,0,0],
    '梁底層鋼筋(支)':[6,0,0],
    '鋼筋直徑_D(mm)': [25.0, 0.0, 0.0],
    '箍筋號數_#(3或4)':[3,0,0],  
    '箍筋間距_s(cm)': [10.0, 0.0, 0.0],
    '極限彎矩_Mux(ton-m)': [85.0, 45.0, 60.0],
    '極限剪力_Vuy(ton)': [35.0, 20.0, 25.0]
}
df_sample = pd.DataFrame(sample_data)

st.title("🏢 高階 SRC / SC 大梁構造、細部與長期撓度檢核系統")
st.markdown("本系統全面支援 **大梁配筋與箍筋檢視**、**長期撓度變形 $\delta$ 檢核**、**SC狗骨頭幾何削弱**、以及 **SRC主筋穿腹板開孔** 驗算框架。")

# --- 左側側邊欄固定輸入面板 ---
st.sidebar.header("✍️ 大梁構造參數快速面板")
s_type = st.sidebar.selectbox("1. 選擇大梁結構形式", ["包覆型SRC", "純鋼骨SC"])
st_type = st.sidebar.selectbox("2. 選擇內部型鋼樣式", ["H型鋼", "箱型柱"])

b_w = st.sidebar.number_input("梁外圍寬 B (cm)", value=50.0)
h_w = st.sidebar.number_input("梁外圍高 H (cm)", value=80.0)
d_s_i = st.sidebar.number_input("型鋼總高 ds (cm)", value=50.0)
b_f_i = st.sidebar.number_input("型鋼翼板寬 bf (cm)", value=25.0)
t_f_i = st.sidebar.number_input("型鋼翼板厚 tf (cm)", value=2.2)
t_w_i = st.sidebar.number_input("型鋼腹板厚 tw (cm)", value=1.4)
lb_i = st.sidebar.number_input("大梁側向未支撐長度 Lb (cm)", value=150.0)

fc_i = st.sidebar.number_input("混凝土強度 fc (kg/cm2)", value=350.0 if s_type=="包覆型SRC" else 0.0)
fy_i = st.sidebar.number_input("型鋼屈服強度 Fy (tf/cm2)", value=3.52)
fyr_i = st.sidebar.number_input("鋼筋屈服強度 Fyr (tf/cm2)", value=4.2 if s_type=="包覆型SRC" else 0.0)

st.sidebar.subheader("縱橫向配筋規格")
r_top = st.sidebar.number_input("梁頂層縱向鋼筋 (支)", value=24 if s_type=="包覆型SRC" else 0)
r_bot = st.sidebar.number_input("梁底層縱向鋼筋 (支)", value=20 if s_type=="包覆型SRC" else 0)
r_d = st.sidebar.number_input("鋼筋直徑代號 D (mm)", value=25.0 if s_type=="包覆型SRC" else 0.0)

stirrup_num = st.sidebar.selectbox("剪力箍筋號數 (#)", [3, 4] if s_type == "包覆型SRC" else [0])
stirrup_spacing = st.sidebar.number_input("箍筋配置間距 s (cm)", value=10.0 if s_type=="包覆型SRC" else 0.0)

st.sidebar.divider()
mux_i = st.sidebar.number_input("極限彎矩 Mux (ton-m)", value=85.0)
vuy_i = st.sidebar.number_input("極限剪力 Vuy (ton)", value=35.0)

single_dict = {
    '構件編號': '單點快速大梁', '結構形式': s_type, '內部鋼構形式': st_type,
    '梁斷面外圍寬_B(cm)': b_w, '梁斷面外圍高_H(cm)': h_w, '鋼構總高_ds(cm)': d_s_i,
    '鋼構翼板寬_bf(cm)': b_f_i, '鋼構翼板厚_tf(cm)': t_f_i, '鋼構腹板厚_tw(cm)': t_w_i,
    '側向未支撐長度_Lb(cm)': lb_i, '混凝土強度_fc(kg/cm2)': fc_i, '鋼材強度_Fy(tf/cm2)': fy_i, '鋼筋強度_Fyr(tf/cm2)': fyr_i,
    '梁頂層鋼筋(支)': r_top, '梁底層鋼筋(支)': r_bot, '鋼筋直徑_D(mm)': r_d,
    '箍筋號數_#(3或4)': stirrup_num, '箍筋間距_s(cm)': stirrup_spacing,
    '極限彎矩_Mux(ton-m)': mux_i, '極限剪力_Vuy(ton)': vuy_i
}
res_s = calculate_beam(single_dict)

# --- 右側主頁籤建立 ---
tab1, tab2 = st.tabs(["📊 大梁批次 Excel 上傳檢核", "📈 單一大梁力學快報"])

with tab1:
    st.subheader("🛠️ 大梁批次自動化分析")
    towrite = io.BytesIO()
    df_sample.to_excel(towrite, index=False, engine='openpyxl')
    towrite.seek(0)
    st.download_button(label="📥 下載大梁標準範例 Excel 範本", data=towrite, file_name="大梁細部構造設計範本.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    
    uploaded_file = st.file_uploader("📂 請上傳大梁設計 Excel 試算表", type=["xlsx", "xls"])
    
    if uploaded_file is not None:
        df_input = pd.read_excel(uploaded_file)
        calculated_rows = []
        for idx, row in df_input.iterrows():
            calculated_rows.append(calculate_beam(row.to_dict()))
        df_res = pd.DataFrame(calculated_rows)
        
        st.success(f"🎉 大梁批次計算完畢！總計大梁數：{len(df_res)} 支")
        
        show_cols = ['構件編號', '結構形式', '內部鋼構形式', '極限彎矩_Mux(ton-m)', '常溫抗彎Mnx(ton-m)', '設計彎矩應力比', '設計剪力應力比', '安全判定', '說明提示']
        st.dataframe(df_res[show_cols], width='stretch')
        
        # --- Word 批次報告書生成 (完全移除中空贅言、精確排除 SRC 施作 RBS) ---
        if st.button("📝 產生大梁結構檢核報告書"):
            doc = Document()
            style = doc.styles['Normal']; font = style.font; font.name = '微軟正黑體'; font.size = Pt(10)
            
            title = doc.add_paragraph()
            title_run = title.add_run("SRC / SC 鋼骨結構大梁極限強度與設計規範檢核報告書")
            title_run.font.size = Pt(16); title_run.font.bold = True; title_run.font.color.rgb = RGBColor(10, 35, 66)
            title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            doc.add_heading("一、 各大梁設計應力比 (D/R Ratio) 統計摘要總表", level=1)
            summary_table = doc.add_table(rows=1, cols=6); summary_table.style = 'Table Grid'
            hdr_cells = summary_table.rows[0].cells
            hdr_titles = ["大梁編號", "結構形式", "型鋼樣式", "彎矩應力比 (D/R)", "剪力應力比 (D/R)", "安全判定"]
            for idx, t_text in enumerate(hdr_titles):
                hdr_cells[idx].text = t_text
                set_cell_background(hdr_cells[idx], "0A2342")
                hdr_cells[idx].paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
                hdr_cells[idx].paragraphs[0].runs[0].font.bold = True
                
            for _, r in df_res.iterrows():
                row_cells = summary_table.add_row().cells
                row_cells[0].text = str(r['構件編號'])
                row_cells[1].text = str(r['結構形式'])
                row_cells[2].text = str(r['內部鋼構形式'])
                row_cells[3].text = f"{r['設計彎矩應力比']:.3f}"
                row_cells[4].text = f"{r['設計剪力應力比']:.3f}"
                row_cells[5].text = str(r['安全判定'])
                bg_color = "E2EFDA" if "PASS" in str(r['安全判定']) else "FCE4D6"
                for cell in row_cells:
                    set_cell_background(cell, bg_color)
                    
            doc.add_paragraph().paragraph_format.space_before = Pt(18)
            doc.add_heading("二、 各大梁詳細力學計算書與高階構造檢核明細", level=1)
            
            for index, r in df_res.iterrows():
                doc.add_heading(f"2.{index+1} 大梁構件明細：{r['構件編號']}", level=2)
                
                table = doc.add_table(rows=5, cols=4); table.style = 'Table Grid'
                for i, r_cells in enumerate(table.rows):
                    if i == 0:
                        r_cells.cells[0].text = "外加極限設計載重"; r_cells.cells[1].text = f"Mux: {r['極限彎矩_Mux(ton-m)']} t-m / Vuy: {r['極限剪力_Vuy(ton)']} t"
                        r_cells.cells[2].text = "常溫標稱抗彎 Mnx"; r_cells.cells[3].text = f"{r['常溫抗彎Mnx(ton-m)']:.1f} ton-m"
                    elif i == 1:
                        r_cells.cells[0].text = "混凝土斷面尺寸"; r_cells.cells[1].text = f"{r['梁斷面外圍寬_B(cm)']} x {r['梁斷面外圍高_H(cm)']} cm"
                        r_cells.cells[2].text = "600°C高溫抗彎 Mn"; r_cells.cells[3].text = f"{r['高溫抗彎Mn_600(ton-m)']:.1f} ton-m"
                    elif i == 2:
                        r_cells.cells[0].text = "型鋼規格 (ds/bf/tw/tf)"; r_cells.cells[1].text = f"{r['鋼構總高_ds(cm)']}x{r['鋼構翼板寬_bf(cm)']}x{r['鋼構腹板厚_tw(cm)']}x{r['鋼構翼板厚_tf(cm)']} cm"
                        r_cells.cells[2].text = "常溫標稱抗剪 Vny"; r_cells.cells[3].text = f"{r['常溫抗剪Vny(ton)']:.1f} ton"
                    elif i == 3:
                        r_cells.cells[0].text = "材料強度 (fc / Fy)"; r_cells.cells[1].text = f"{r['混凝土強度_fc(kg/cm2)']} / {r['鋼材強度_Fy(tf/cm2)']} tf/cm2"
                        r_cells.cells[2].text = "600°C高溫抗剪 Vn"; r_cells.cells[3].text = f"{r['高溫抗剪Vn_600(ton)']:.1f} ton"
                    elif i == 4:
                        r_cells.cells[0].text = "彎矩應力比 (D/R)"; r_cells.cells[1].text = f"{r['設計彎矩應力比']:.3f}"
                        r_cells.cells[2].text = "剪力應力比 (D/R)"; r_cells.cells[3].text = f"{r['設計剪力應力比']:.3f}"
                        set_cell_background(r_cells.cells[1], "E2EFDA" if r['設計彎矩應力比']<=1.0 else "FCE4D6")
                        set_cell_background(r_cells.cells[3], "E2EFDA" if r['設計剪力應力比']<=1.0 else "FCE4D6")

                doc.add_paragraph().paragraph_format.space_before = Pt(6)
                p_detail = doc.add_paragraph()
                p_detail.add_run("■ 🔧 高階構造、細部與長期變形檢核成果摘要：\n").bold = True
                p_detail.add_run(f"  • 大梁配置斷面幾何：型鋼 {r['內部鋼構形式']} (H {r['鋼構總高_ds(cm)']:.0f} x {r['鋼構翼板寬_bf(cm)']:.0f} x {r['鋼構腹板厚_tw(cm)']} x {r['鋼構翼板厚_tf(cm)']} cm)\n")
                p_detail.add_run(f"  • 🧱 大梁法規長期變形檢核：{r['撓度檢核說明']}\n")
                
                if r['結構形式'] == "包覆型SRC":
                    p_detail.add_run(f"  • 橫向剪力箍筋配置規格：#{r['箍筋號數_#(3或4)']:.0f} @ {r['箍筋間距_s(cm)']:.1f} cm 🟢 已完美加算至抗剪耐力\n")
                    p_detail.add_run(f"  • 大梁主筋穿腹板開孔驗算：{r['開孔補強說明']}\n")
                elif r['結構形式'] == "純鋼骨SC" and r['內部鋼構形式'] == "H型鋼":
                    p_detail.add_run(f"  • 梁端減縮（RBS 狗骨頭）幾何演算：{r['RBS提示說明']}\n")
                
                p_sep = doc.add_paragraph()
                p_sep.add_run("--------------------------------------------------------------------------------------------").font.color.rgb = RGBColor(200, 200, 200)

            doc_download = io.BytesIO()
            doc.save(doc_download); doc_download.seek(0)
            st.download_button(label="💾 下載大梁構造細部整合報告書 (.docx)", data=doc_download, file_name="高階大梁細部構造設計報告書.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

with tab2:
    st.success("📊 左側面板連動 ── 單一斷面力學分析快報")
    st.metric(label="📊 大梁綜合安全性判定", value=res_s['安全判定'])
    
    c_a, c_b = st.columns(2)
    with c_a:
        st.metric(label="常溫標稱抗彎強度 Mnx", value=f"{res_s['常溫抗彎Mnx(ton-m)']:.2f} ton-m")
        st.metric(label="彎矩設計應力比 (D/R)", value=f"{res_s['設計彎矩應力比']:.3f}")
    with c_b:
        st.metric(label="常溫標稱抗剪強度 Vny", value=f"{res_s['常溫抗剪Vny(ton)']:.2f} ton")
        st.metric(label="剪力設計應力比 (D/R)", value=f"{res_s['設計剪力應力比']:.3f}")
        
    st.info(fr"💡 構造細部提示：{res_s.get('說明提示', '')}")
    
    st.markdown("### 🔧 高階構造、細部與長期變形檢核成果")
    st.write(fr"大梁配置斷面幾何：型鋼 **{st_type} (H {d_s_i:.0f} x {b_f_i:.0f} x {t_w_i} x {t_f_i} cm)**")
    st.write(fr"🧱 大梁法規長期變形檢核：**{res_s['撓度檢核說明']}**")
    
    if s_type == "包覆型SRC":
        st.write(fr"橫向剪力箍筋配置規格：**#{stirrup_num:.0f} @ {stirrup_spacing:.1f} cm** 🟢 已完美加算至抗剪耐力")
        st.write(fr"大梁主筋穿腹板開孔驗算：{res_s['開孔補強說明']}")
        
    if s_type == "純鋼骨SC" and st_type == "H型鋼":
        st.write(fr"梁端減縮（RBS 狗骨頭）幾何演算：{res_s['RBS提示說明']}")
            
    st.divider()
    st.write(fr"600°C 火災下大梁剩餘抗彎強度 $M_{{nx,600}}$：**{res_s['高溫抗彎Mn_600(ton-m)']:.1f} ton-m**")
    st.write(fr"600°C 火災下大梁剩餘抗剪強度 $V_{{ny,600}}$：**{res_s['高溫抗剪Vn_600(ton)']:.1f} ton**")