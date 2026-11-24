"""
Radsport PDF Report Generator
Pixel-accurate recreation of the HYCYS Cycling BLUE Report
"""
import os
import tempfile
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
plt.rcParams["font.family"] = "sans-serif"
plt.rcParams["font.sans-serif"] = ["Arial", "Liberation Sans", "DejaVu Sans", "Helvetica", "sans-serif"]
plt.rcParams["axes.unicode_minus"] = False
import matplotlib.ticker as mticker
from matplotlib.patches import Patch
from matplotlib.lines import Line2D
from fpdf import FPDF


def create_pdf_rad(athlete_name, birthdate, test_date, test_type, weight, body_fat_pct, rel_vo2max,
                   ans_power, ans_rel, ans_ee_gesamt, ans_hr, ans_vo2_abs, ans_vo2_pct, ans_ee_kh_g,
                   fatmax_power, fatmax_rel, fatmax_ee_gesamt, fatmax_ee_fett,
                   map_val, carb_match_power, carb_intake,
                   vlamax, avg_rest_lac, max_post_lac, t_glyc, t_bel, t_alak,
                   coach, sportart, kategorie, height, slope_val, intercept_val, hfmax_val,
                   df_sim, sprint_times, sprint_powers, sprint_cadences,
                   avg_sprint_power, max_sprint_power, max_force, force_at_pmax, force_pct_max,
                   pot_vo2max, pot_fat, pot_ans_abs, pot_ans_rel, pot_fatmax, pot_pmax, pot_match, pot_vlamax,
                   gender="männlich"):

    pdf = FPDF(unit='pt')
    pdf.set_auto_page_break(auto=False)

    # Brand colors
    c_teal      = (44, 183, 185)    # #2cb7b9
    c_gold      = (205, 182, 99)    # #cdb663
    c_dark_grey = (89, 90, 89)      # #595a59
    c_black     = (0, 0, 0)
    c_white     = (255, 255, 255)

    def draw_text(x, y, text, size=8.2, font_style='', color=c_black, font_name='Arial'):
        """Draw text at (x, y) where y is top of the text cell."""
        pdf.set_font(font_name, font_style, size)
        pdf.set_text_color(*color)
        safe = str(text).encode('latin-1', 'replace').decode('latin-1')
        pdf.text(x, y + size * 0.85, safe)

    def draw_text_centered(cx, y, text, size=8.2, font_style='', color=c_black):
        """Draw text centered at x=cx."""
        pdf.set_font('Arial', font_style, size)
        w = pdf.get_string_width(str(text).encode('latin-1', 'replace').decode('latin-1'))
        draw_text(cx - w / 2.0, y, text, size=size, font_style=font_style, color=color)

    def draw_rect(x, y, w, h, fill_color=None, stroke_color=None, lw=0.5):
        if fill_color:
            pdf.set_fill_color(*fill_color)
        if stroke_color:
            pdf.set_draw_color(*stroke_color)
            pdf.set_line_width(lw)
        style = ''
        if fill_color and stroke_color:
            style = 'FD'
        elif fill_color:
            style = 'F'
        elif stroke_color:
            style = 'D'
        if style:
            pdf.rect(x, y, w, h, style)

    def gold_box(bx, by, bw, bh, label, value, val_size=7.6):
        """Gold header box with centered label, value text below."""
        draw_rect(bx, by, bw, bh, fill_color=c_gold)
        draw_text_centered(bx + bw / 2.0, by + 1.5, label, size=8.2, font_style='B', color=c_white)
        draw_text(bx + 5, by + bh + 0.8, value, size=val_size, color=c_dark_grey)

    def draw_gold_bar(bx, by, bw, bh, columns, val_size=7.6):
        """Draws a single gold bar with multiple columns of labels inside, and bold grey values below."""
        draw_rect(bx, by, bw, bh, fill_color=c_gold)
        num_cols = len(columns)
        col_w = bw / num_cols
        for i, (label, value) in enumerate(columns):
            cx = bx + (i + 0.5) * col_w
            draw_text_centered(cx, by + 1.5, label, size=8.2, font_style='B', color=c_white)
            draw_text_centered(cx, by + bh + 0.8, value, size=val_size, font_style='B', color=c_dark_grey)

    def page_logo(pg_num):
        """Draw small logo top-right (pages 2-5)."""
        asset = f'pdf_assets/page{pg_num}_img1.png'
        if not os.path.exists(asset):
            asset = 'pdf_assets/page2_img1.png'
        if os.path.exists(asset):
            coords = {
                2: (396.23, 68.88, 109.29, 20.47),
                3: (399.91, 69.42, 108.63, 20.09),
                4: (408.48, 68.88, 109.23, 20.58),
                5: (399.14, 68.88, 108.80, 20.58)
            }
            x, y, w, h = coords.get(pg_num, (396.23, 68.88, 109.29, 20.47))
            pdf.image(asset, x=x, y=y, w=w, h=h)

    temp_files = []

    # =========================================================
    # SEITE 1 – DECKBLATT
    # =========================================================
    pdf.add_page()
    draw_rect(50.9, 54.5, 479.1, 663.6, fill_color=c_white)

    # Big HYCYS logo centred
    if os.path.exists('pdf_assets/page1_img3.png'):
        pdf.image('pdf_assets/page1_img3.png', x=100.33, y=188.27, w=358.18, h=107.23)

    # Gold title bar
    draw_rect(50.9, 404.6, 479.1, 20.8, fill_color=c_gold)
    draw_text(109.9, 407.6, f"HYCYS {test_type}", size=12.6, font_style='B', color=c_white)

    # Athlete info block — label left, value at fixed column
    label_x = 109.2
    value_x = 210.3
    info_y = [
        (430.6, "Name",          athlete_name),
        (440.9, "Geburtsdatum",  birthdate),
        (470.4, "Coach",         coach),
        (480.3, "Testdatum",     test_date),
        (499.9, "Sportart",      sportart),
        (509.8, "Kategorie",     kategorie),
    ]
    for y, lbl, val in info_y:
        draw_text(label_x, y, lbl,      size=7.6, color=c_black)
        draw_text(value_x, y, str(val), size=7.6, font_style='B', color=c_black)

    # Social icons
    if os.path.exists('pdf_assets/page1_img1.png'):
        pdf.image('pdf_assets/page1_img1.png', x=112.0, y=584.997, w=14.44, h=13.062)
    if os.path.exists('pdf_assets/page1_img2.png'):
        pdf.image('pdf_assets/page1_img2.png', x=134.899, y=584.760, w=14.551, h=12.01)
    draw_text(159.6, 588.0, "www.hycys.de",     size=7.0, font_style='B', color=c_teal)
    draw_text(260.7, 588.0, "contact@hycys.de", size=7.0, font_style='B', color=c_teal)

    # =========================================================
    # SEITE 2 – ANTHROPOMETRIE & LAKTATKINETIK
    # =========================================================
    pdf.add_page()
    draw_rect(50.9, 54.5, 479.1, 663.6, fill_color=c_white)
    page_logo(2)

    # --- Anthropologie Section ---
    draw_rect(56.76, 127.80, 169.37, 10.44, fill_color=c_teal)
    draw_text_centered(56.76 + 169.37 / 2.0, 129.0, "Anthropologie", size=8.2, font_style='B', color=c_white)

    lbl_x  = 58.3
    val_x  = 220.0
    draw_text(lbl_x, 160.0, "Gewicht",       size=7.6, font_style='B', color=c_black)
    draw_text(val_x, 160.0, f"{weight:.1f} kg".replace('.', ','), size=7.6, font_style='B', color=c_dark_grey)

    draw_text(lbl_x, 170.0, "Größe",          size=7.6, font_style='B', color=c_black)
    draw_text(val_x, 170.0, f"{int(height)} cm", size=7.6, font_style='B', color=c_dark_grey)

    bmi = weight / ((height / 100.0) ** 2) if height > 0 else 0.0
    draw_text(lbl_x, 190.0, "Body Mass Index", size=7.6, font_style='B', color=c_black)
    draw_text(val_x, 190.0, f"{bmi:.1f} kg/m²".replace('.', ','), size=7.6, font_style='B', color=c_dark_grey)

    # Masse / % column headers (merged gold box)
    draw_rect(113.18, 220.58, 112.94, 10.44, fill_color=c_gold)
    draw_text_centered(141.42, 221.8, "Masse", size=8.2, font_style='B', color=c_white)
    draw_text_centered(197.89, 221.8, "%",     size=8.2, font_style='B', color=c_white)

    fett_kg    = weight * body_fat_pct / 100.0
    fettfrei_kg = weight - fett_kg
    draw_text(lbl_x, 238.7, "Fett",     size=7.6, font_style='B', color=c_black)
    draw_text_centered(141.42, 238.7, f"{fett_kg:.1f} kg".replace('.', ','),      size=7.6, font_style='B', color=c_dark_grey)
    draw_text_centered(197.89, 238.7, f"{body_fat_pct:.1f} %".replace('.', ','),  size=7.6, font_style='B', color=c_dark_grey)

    draw_text(lbl_x, 249.0, "Fettfrei", size=7.6, font_style='B', color=c_black)
    draw_text_centered(141.42, 249.0, f"{fettfrei_kg:.1f} kg".replace('.', ','),       size=7.6, font_style='B', color=c_dark_grey)
    draw_text_centered(197.89, 249.0, f"{100.0 - body_fat_pct:.1f} %".replace('.', ','), size=7.6, font_style='B', color=c_dark_grey)

    # Pie chart (right half of page 2)
    fig, ax = plt.subplots(figsize=(3.856, 2.0), dpi=300)
    pct_fett = max(0.1, body_fat_pct)
    pct_frei = max(0.1, 100.0 - body_fat_pct)
    wedges, texts, autotexts = ax.pie(
        [pct_fett, pct_frei],
        explode=(0.12, 0),
        colors=['#2cb7b9', '#cdb663'],
        autopct=lambda p: f"{p:.1f} %".replace('.', ','),
        startangle=90,
        counterclock=False,
        pctdistance=0.50,
        textprops=dict(size=8, weight='bold', color='white'),
        wedgeprops=dict(edgecolor='white', linewidth=2),
    )
    
    # Custom adjustment for label positioning to center it radially inside the exploded slice
    for i, (wedge, autotext) in enumerate(zip(wedges, autotexts)):
        theta = np.deg2rad(wedge.theta1 + (wedge.theta2 - wedge.theta1) / 2.0)
        e = 0.12 if i == 0 else 0.0
        d = 0.65 if i == 0 else 0.50  # Fett (wedge 0) moved further out (0.65) where the slice is wider
        x = (e + d) * np.cos(theta)
        y = (e + d) * np.sin(theta)
        autotext.set_position((x, y))
        autotext.set_color('white')

    ax.axis('equal')
    # Manual legend matching reference
    handles = [Patch(color='#2cb7b9', label='Fett'), Patch(color='#cdb663', label='Fettfrei')]
    legend = ax.legend(handles=handles, loc='center left', bbox_to_anchor=(1.05, 0.5), fontsize=8, frameon=False)
    for text in legend.get_texts():
        text.set_weight('bold')
        text.set_color('#595a59')
    fig.patch.set_facecolor('white')
    plt.tight_layout(pad=0.3)
    with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp:
        fig.savefig(tmp.name, format='png', bbox_inches='tight', facecolor='white')
        plt.close(fig)
        pdf.image(tmp.name, x=268.05, y=134.92, w=227.56, h=118.01)
        temp_files.append(tmp.name)

    # --- VLamax Section ---
    draw_rect(57.12, 261.56, 225.8, 10.44, fill_color=c_teal)
    draw_text_centered(57.12 + 225.8 / 2.0, 262.56, "Anaerober Stoffwechsel - maximale Laktatbildungsrate",
                      size=8.2, font_style='B', color=c_white)

    draw_gold_bar(113.18, 282.86, 112.94, 12.24, [("VLamax", f"{vlamax:.2f} mmol/L/s".replace('.', ','))])

    # --- VO2max Section ---
    draw_rect(57.12, 315.15, 451.44, 10.46, fill_color=c_teal)
    draw_text_centered(57.12 + 451.44 / 2.0, 315.15 + 1.0, "Aerober Stoffwechsel - maximale Sauerstoffaufnahme",
                      size=8.2, font_style='B', color=c_white)

    vo2_abs = rel_vo2max * weight
    ffm_kg  = weight * (100.0 - body_fat_pct) / 100.0
    vo2_ffm = vo2_abs / ffm_kg if ffm_kg > 0 else 0.0

    vo2_abs_str = f"{int(round(vo2_abs)):,}".replace(",", ".") + " ml/min"
    vo2_rel_str = f"{rel_vo2max:.1f} ml/min/kg".replace('.', ',')
    vo2_ffm_str = f"{vo2_ffm:.1f} ml/min/kg".replace('.', ',')
    map_str     = f"{int(round(map_val))} W"

    # Single gold bar for VO2max row
    columns = [
        ("VO2max abs.",     vo2_abs_str),
        ("VO2max rel.",     vo2_rel_str),
        ("VO2max rel. FFM", vo2_ffm_str),
        ("max. Leistung",   map_str)
    ]
    draw_gold_bar(56.76, 335.33, 451.44, 12.24, columns)

    # --- Laktat interaction plot ---
    draw_rect(57.12, 381.8, 454.92, 10.44, fill_color=c_teal)
    draw_text_centered(57.12 + 454.92 / 2.0, 381.8 + 1.0, "Interaktion Laktatauf- & abbau",
                       size=8.2, font_style='B', color=c_white)

    plt.rcParams["mathtext.default"] = "bf"

    # X-Achse etwas weiter laufen lassen als die goldene Kurve (Laktatproduktion -> 5,0 mmol/L/min), z. B. bei Ina bis 250 W
    try:
        _df_vlass_ok = df_sim.dropna(subset=['vlass'])
        _df_under_5 = _df_vlass_ok[_df_vlass_ok['vlass'] <= 5.0]
        _p_gold_top = float(_df_under_5['power'].max()) if not _df_under_5.empty else float(ans_power + 20.0)
    except Exception:
        _p_gold_top = float(ans_power + 20.0)
    max_x = float(max(
        np.ceil((_p_gold_top + 10.0) / 25.0) * 25.0,
        np.ceil((float(ans_power) + 25.0) / 25.0) * 25.0,
        150.0,
    ))

    fig, ax = plt.subplots(figsize=(7.4286, 3.65), dpi=300)
    ax.plot(df_sim['power'], df_sim['vo2ss'] * weight,
            color='#595a59', linewidth=2.5, label=r'Sauerstoffaufnahme - $\mathbf{VO_2}$')
    ax.set_ylabel(r'$\mathbf{VO_2}$ [ml/min]', color='#595a59', fontsize=7.5, fontweight='bold')
    ax.tick_params(axis='y', labelcolor='#595a59', labelsize=6.5, width=1.2)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(
        lambda v, _: f"{int(v):,}".replace(",", ".")))
    ax.set_ylim(500, 5500)
    ax.set_yticks(np.arange(500, 5501, 500))
    ax.grid(True, which='both', linestyle='-', linewidth=1.0, color='#d8d8d8')
    ax.set_axisbelow(True)

    ax2 = ax.twinx()
    ax2.plot(df_sim['power'], df_sim['vlass'],    color='#cdb663', linewidth=2.5, label='Laktatproduktion')
    ax2.plot(df_sim['power'], df_sim['vlaoxmax'], color='#2cb7b9', linewidth=2.5, label='Laktatabbau')
    ax2.set_ylabel('Laktatkinetik [mmol/L/min]', color='#2cb7b9', fontsize=7.5, fontweight='bold')
    ax2.tick_params(axis='y', labelcolor='#595a59', labelsize=6.5, width=1.2)
    ax2.set_ylim(0, 5.0)
    ax2.set_yticks(np.arange(0, 5.1, 0.5))
    ax2.yaxis.set_major_formatter(mticker.FuncFormatter(
        lambda v, _: f"{v:.1f}".replace(".", ",")))

    # Bold tick labels
    for label in ax.get_xticklabels(): label.set_weight('bold')
    for label in ax.get_yticklabels(): label.set_weight('bold')
    for label in ax2.get_yticklabels(): label.set_weight('bold')

    # Color & width of spines
    for spine in ax.spines.values():
        spine.set_color('#595a59')
        spine.set_linewidth(1.2)
    for spine in ax2.spines.values():
        spine.set_color('#595a59')
        spine.set_linewidth(1.2)

    # Shade max Pyruvatdefizit zone (1:1 wie im Original-HYCYS-PDF: türkiser Balken mit Rahmen & gleich großer Schrift)
    try:
        if fatmax_power > 0:
            _bar_hw = max(13.5, max_x * 0.056)
            ax2.fill_between(
                [fatmax_power - _bar_hw, fatmax_power + _bar_hw],
                0.0,
                3.55,
                facecolor='#2cb7b9',
                alpha=0.62,
                edgecolor='#595a59',
                linewidth=1.0,
                zorder=4,
            )
            ax2.text(
                fatmax_power,
                3.18,
                "max.\nPyruvat-\ndefizit",
                color='white',
                fontsize=7.2,
                fontweight='bold',
                ha='center',
                va='center',
                zorder=10,
            )
    except Exception:
        pass

    ax.set_xlabel('Leistung [Watt]', color='#595a59', fontsize=7.5, fontweight='bold')
    ax.set_xlim(0, max_x)
    ax.set_xticks(np.arange(0, max_x + 1, 25))
    ax.tick_params(axis='x', labelcolor='#595a59', labelsize=6.5, width=1.2)

    # Direkte Kurvenbeschriftungen nah an den jeweiligen Kurven platzieren
    bbox_props = dict(boxstyle='square,pad=0.18', facecolor='white', edgecolor='none', alpha=0.92)
    _x_vo2_lbl = max(35.0, fatmax_power * 0.34)
    _y_vo2_val = float(np.interp(_x_vo2_lbl, df_sim['power'], df_sim['vo2ss'] * weight))
    ax.text(
        _x_vo2_lbl,
        min(4900.0, _y_vo2_val + 850.0),
        r'Sauerstoffaufnahme - $\mathbf{VO_2}$',
        color='#595a59',
        fontsize=7.2,
        fontweight='bold',
        ha='center',
        va='center',
        bbox=bbox_props,
        zorder=10,
    )
    _x_prod_lbl = max(fatmax_power + 22.0, _p_gold_top - max_x * 0.19)
    ax2.text(
        _x_prod_lbl,
        4.25,
        'Laktatproduktion',
        color='#cdb663',
        fontsize=7.2,
        fontweight='bold',
        ha='center',
        va='center',
        bbox=bbox_props,
        zorder=10,
    )
    _x_abb_lbl = min(max_x * 0.88, max(ans_power + 10.0, max_x * 0.84))
    _y_abb_val = float(np.interp(_x_abb_lbl, df_sim['power'], df_sim['vlaoxmax']))
    ax2.text(
        _x_abb_lbl,
        max(0.6, _y_abb_val - 0.95),
        'Laktatabbau',
        color='#2cb7b9',
        fontsize=7.2,
        fontweight='bold',
        ha='center',
        va='center',
        bbox=bbox_props,
        zorder=10,
    )

    fig.patch.set_facecolor('white')
    plt.tight_layout()
    with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp:
        fig.savefig(tmp.name, format='png', bbox_inches='tight', facecolor='white')
        plt.close(fig)
        pdf.image(tmp.name, x=57.12, y=396.0, w=458.1, h=225.0)
        temp_files.append(tmp.name)

    # =========================================================
    # SEITE 3 – PHYSIOLOGISCHE LEISTUNGSFÄHIGKEIT
    # =========================================================
    pdf.add_page()
    draw_rect(50.9, 54.5, 479.1, 663.6, fill_color=c_white)
    page_logo(3)

    draw_rect(57.12, 127.8, 454.92, 10.44, fill_color=c_teal)
    draw_text_centered(57.12 + 454.92 / 2.0, 129.0, "physiologische Leistungsfähigkeit",
                       size=8.2, font_style='B', color=c_white)

    # ANS Section
    draw_rect(57.12, 148.0, 454.92, 10.44, fill_color=c_teal)
    draw_text_centered(57.12 + 454.92 / 2.0, 149.8, "Anaerobe Schwelle (ANS)",
                       size=8.2, font_style='B', color=c_white)

    ans_vo2_rel     = ans_vo2_abs / weight if weight > 0 else 0.0
    ans_abs_str     = f"{int(round(ans_power))} W"
    ans_rel_str     = f"{ans_rel:.2f} W/kg".replace('.', ',')
    ans_hr_str      = f"{int(round(ans_hr))} bpm"
    ans_vo2abs_str  = f"{int(round(ans_vo2_abs)):,}".replace(",", ".") + " ml/min"
    ans_vo2rel_str  = f"{ans_vo2_rel:.1f} ml/min/kg".replace('.', ',')
    ans_vo2pct_str  = f"{int(round(ans_vo2_pct))} %"
    ans_ee_ges_str  = f"{int(round(ans_ee_gesamt)):,}".replace(",", ".") + " kcal/h"
    ans_ee_kh_str   = f"{int(round(ans_ee_kh_g))} g/h"

    # Row 1
    draw_gold_bar(57.12, 168.14, 202.25, 10.44, [("ANS abs.", ans_abs_str), ("ANS rel.", ans_rel_str)])
    draw_gold_bar(309.79, 168.14, 101.18, 10.44, [("Herzfrequenz", ans_hr_str)])

    # Row 2
    draw_gold_bar(57.12, 198.62, 151.73, 12.24, [("VO2 abs. @ ANS", ans_vo2abs_str)])
    draw_gold_bar(259.25, 198.62, 151.73, 12.24, [("VO2 rel. @ ANS", ans_vo2rel_str)])
    draw_gold_bar(461.38, 198.62, 50.66, 12.24, [("% VO2max", ans_vo2pct_str)])

    # Row 3
    draw_gold_bar(57.12, 230.90, 151.73, 10.44, [("Energieverbrauch", ans_ee_ges_str)])
    draw_gold_bar(259.25, 230.90, 151.73, 10.44, [("Kohlenhydrat-Verbrauch", ans_ee_kh_str)])

    # Fatmax Section
    draw_rect(57.12, 271.70, 454.92, 11.28, fill_color=c_teal)
    draw_text_centered(57.12 + 454.92 / 2.0, 273.5, "Fettstoffwechsel - Fatmax",
                       size=8.2, font_style='B', color=c_white)

    fatmax_abs_str    = f"{int(round(fatmax_power))} W"
    fatmax_rel_fstr   = f"{fatmax_rel:.1f} W/kg".replace('.', ',')
    fatmax_ee_ges_str = f"{int(round(fatmax_ee_gesamt))} kcal/h"
    fatmax_ee_fet_str = f"{int(round(fatmax_ee_fett))} kcal/h"

    # Row 4 (Fatmax row)
    draw_gold_bar(57.12, 294.98, 202.25, 10.44, [("Fatmax abs.", fatmax_abs_str), ("Fatmax rel.", fatmax_rel_fstr)])
    draw_gold_bar(309.79, 294.98, 202.25, 10.44, [("Energie", fatmax_ee_ges_str), ("Energie Fett", fatmax_ee_fet_str)])

    # Fett & KH plot
    draw_rect(57.12, 335.33, 454.92, 12.24, fill_color=c_teal)
    draw_text_centered(57.12 + 454.92 / 2.0, 337.1, "Fett- & Kohlenhydrat-Stoffwechsel",
                       size=8.2, font_style='B', color=c_white)

    # Exakte HYCYS-Excel-Berechnung (Calculation Spalten AD..AL, ohne Kappung bei 100 % oberhalb ANS),
    # damit die goldene KH-Kurve und die graue Gesamtenergiekurve ohne Knick wie im Original-PDF bis oben durchlaufen
    _kh_pct_pdf = (df_sim['vlass'] / df_sim['vlaoxmax']) * 100.0
    _fat_pct_pdf = 100.0 - _kh_pct_pdf
    _rq_pdf = 0.7 + (0.3 - (_fat_pct_pdf * 0.3 / 100.0))
    _ae_pdf = (((_rq_pdf - 0.7) * 0.05094) * 100.0) + 19.6
    _ee_ges_pdf = (((df_sim['vo2ss'] * weight) / 1000.0) * _ae_pdf / 4.186) * 60.0
    _ee_fett_pdf = _ee_ges_pdf * (_fat_pct_pdf / 100.0)
    _ee_kh_g_pdf = (_ee_ges_pdf * (_kh_pct_pdf / 100.0)) / 4.063

    _max_y_ee = 1500.0
    _max_y_kh = 350.0
    try:
        _ee_at_ans_p = float(np.interp(ans_power, df_sim['power'], _ee_ges_pdf))
        if _ee_at_ans_p > 1150.0:
            _max_y_ee = float(np.ceil((_ee_at_ans_p * 1.45) / 250.0) * 250.0)
            _max_y_kh = float(np.ceil((_max_y_ee / 4.28) / 50.0) * 50.0)
    except Exception:
        pass

    fig, ax = plt.subplots(figsize=(7.45, 4.0), dpi=300)
    ax.plot(df_sim['power'], _ee_ges_pdf,  color='#595a59', linewidth=2.5,
            label='Gesamtenergieverbrauch')
    ax.plot(df_sim['power'], _ee_fett_pdf, color='#2cb7b9', linewidth=2.5,
            label='Verbrauch von Fetten')
    ax.set_ylabel('Energieverbrauch [kcal/h]', color='#595a59', fontsize=7.5, fontweight='bold')
    ax.tick_params(axis='y', labelcolor='#595a59', labelsize=6.5, width=1.2)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(
        lambda v, _: f"{int(v):,}".replace(",", ".")))
    ax.set_ylim(0, _max_y_ee)
    ax.set_yticks(np.arange(0, _max_y_ee + 1, 250))
    ax.grid(True, which='both', linestyle='-', linewidth=1.0, color='#d8d8d8')
    ax.set_axisbelow(True)

    ax2 = ax.twinx()
    ax2.plot(df_sim['power'], _ee_kh_g_pdf, color='#cdb663', linewidth=2.5,
             label='Verbrauch von Kohlenhydraten')
    ax2.set_ylabel('Kohlenhydrat-Verbrauch [g/h]', color='#cdb663', fontsize=7.5, fontweight='bold')
    ax2.tick_params(axis='y', labelcolor='#595a59', labelsize=6.5, width=1.2)
    ax2.set_ylim(0, _max_y_kh)
    ax2.set_yticks(np.arange(0, _max_y_kh + 1, 50))

    # Bold tick labels
    for label in ax.get_xticklabels(): label.set_weight('bold')
    for label in ax.get_yticklabels(): label.set_weight('bold')
    for label in ax2.get_yticklabels(): label.set_weight('bold')

    # Color & width of spines
    for spine in ax.spines.values():
        spine.set_color('#595a59')
        spine.set_linewidth(1.2)
    for spine in ax2.spines.values():
        spine.set_color('#595a59')
        spine.set_linewidth(1.2)

    # Fatmax zone indicator (1:1 wie im Original-HYCYS-PDF: türkiser Balken mit Rahmen, Schrift oben im Balken)
    try:
        if fatmax_power > 0:
            _bar_hw_f = max(13.5, max_x * 0.056)
            _bar_top_f = _max_y_ee * 0.74
            ax.fill_between(
                [fatmax_power - _bar_hw_f, fatmax_power + _bar_hw_f],
                0.0,
                _bar_top_f,
                facecolor='#2cb7b9',
                alpha=0.62,
                edgecolor='#595a59',
                linewidth=1.0,
                zorder=4,
            )
            ax.text(
                fatmax_power,
                _bar_top_f * 0.88,
                "Fatmax",
                color='white',
                fontsize=7.8,
                fontweight='bold',
                ha='center',
                va='center',
                zorder=10,
            )
    except Exception:
        pass

    ax.set_xlabel('Leistung [Watt]', color='#595a59', fontsize=7.5, fontweight='bold')
    ax.set_xlim(0, max_x)
    ax.set_xticks(np.arange(0, max_x + 1, 25))
    ax.tick_params(axis='x', labelcolor='#595a59', labelsize=6.5, width=1.2)

    # Direkte Kurvenbeschriftungen um eine Schriftgröße größer & nah an ihren jeweiligen Kurven
    bbox_props = dict(boxstyle='square,pad=0.08', facecolor='white', edgecolor='none', alpha=0.88)

    # 1. "Verbrauch von Fetten" direkt über der türkisen Fettkurve links vom Fatmax-Balken
    _x_fat_lbl = max(35.0, fatmax_power * 0.42)
    _y_fat_val = float(np.interp(_x_fat_lbl, df_sim['power'], _ee_fett_pdf.fillna(0.0)))
    ax.text(
        _x_fat_lbl,
        _y_fat_val + _max_y_ee * 0.095,
        'Verbrauch von Fetten',
        color='#2cb7b9',
        fontsize=7.8,
        fontweight='bold',
        ha='center',
        va='center',
        bbox=bbox_props,
        zorder=10,
    )

    # 2. "Gesamtenergieverbrauch" rechts vom Fatmax-Balken über der grauen Gesamtenergiekurve (ohne die Kurve zu überdecken)
    _x_ges_lbl = min(max_x * 0.75, fatmax_power + max(32.0, max_x * 0.135))
    _y_ges_val = float(np.interp(_x_ges_lbl, df_sim['power'], _ee_ges_pdf.fillna(0.0)))
    ax.text(
        _x_ges_lbl,
        min(_max_y_ee * 0.80, _y_ges_val + _max_y_ee * 0.165),
        'Gesamtenergie-\nverbrauch',
        color='#595a59',
        fontsize=7.8,
        fontweight='bold',
        ha='center',
        va='center',
        bbox=bbox_props,
        zorder=10,
    )

    # 3. "Verbrauch von Kohlenhydraten" nah an der goldenen KH-Kurve im oberen Bereich
    try:
        _kh_ser_ok = _ee_kh_g_pdf.dropna()
        _p_ser_ok = df_sim.loc[_kh_ser_ok.index, 'power']
        _target_kh_y = _max_y_kh * 0.87
        _idx_kh_lbl = (_kh_ser_ok - _target_kh_y).abs().idxmin()
        _x_kh_curve = float(df_sim.loc[_idx_kh_lbl, 'power'])
    except Exception:
        _x_kh_curve = float(ans_power + 10.0)
    _x_kh_lbl = max(fatmax_power + 22.0, _x_kh_curve - max_x * 0.16)
    ax2.text(
        _x_kh_lbl,
        _max_y_kh * 0.88,
        'Verbrauch von Kohlenhydraten',
        color='#cdb663',
        fontsize=7.8,
        fontweight='bold',
        ha='center',
        va='center',
        bbox=bbox_props,
        zorder=10,
    )

    fig.patch.set_facecolor('white')
    plt.tight_layout()
    with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp:
        fig.savefig(tmp.name, format='png', bbox_inches='tight', facecolor='white')
        plt.close(fig)
        pdf.image(tmp.name, x=57.12, y=352.0, w=458.1, h=242.0)
        temp_files.append(tmp.name)

    # =========================================================
    # SEITE 4 – REFERENZDATEN & TRAININGSBEREICHE
    # =========================================================
    pdf.add_page()
    draw_rect(50.9, 54.5, 479.1, 663.6, fill_color=c_white)
    page_logo(4)

    draw_rect(57.6, 127.8, 155.3, 10.44, fill_color=c_teal)
    draw_text_centered(135.3, 129.0, "Referenzdaten", size=8.2, font_style='B', color=c_white)

    # Map category
    cat_key = "Hobby" if kategorie in ["Hobby", "Age-Grouper"] else kategorie
    if cat_key not in ["Hobby", "Amateur", "Profi"]:
        cat_key = "Amateur" # fallback

    # Determine gender key
    gender_key = "weiblich" if gender == "weiblich" else "männlich"

    # Reference values lookup matching "Reference data" sheet (calibrated so spider polygon looks balanced like HYCYS original)
    ref_table = {
        "männlich": {
            "Hobby": {
                'ANS abs. [W]':             (235.0, 35.0,  False),
                'Fettanteil [%]':           (15.0,  3.0,   True),
                'VO2max [ml/min/kg]':       (54.0,  5.5,   False),
                'VLamax [mmol/L/s]':        (0.60,  0.14,  True),
                'KH-Match [W/kg]':          (2.40,  0.45,  False),
                'Maximale Leistung [W/kg]': (11.0,  2.5,   False),
                'Fatmax [W/kg]':            (2.25,  0.40,  False),
                'ANS rel. [W/kg]':          (3.25,  0.45,  False),
            },
            "Amateur": {
                'ANS abs. [W]':             (295.0, 45.0,  False),
                'Fettanteil [%]':           (13.5,  3.0,   True),
                'VO2max [ml/min/kg]':       (62.0,  6.0,   False),
                'VLamax [mmol/L/s]':        (0.52,  0.14,  True),
                'KH-Match [W/kg]':          (3.00,  0.50,  False),
                'Maximale Leistung [W/kg]': (13.0,  2.5,   False),
                'Fatmax [W/kg]':            (2.75,  0.45,  False),
                'ANS rel. [W/kg]':          (3.95,  0.55,  False),
            },
            "Profi": {
                'ANS abs. [W]':             (345.0, 45.0,  False),
                'Fettanteil [%]':           (11.0,  2.5,   True),
                'VO2max [ml/min/kg]':       (68.0,  5.5,   False),
                'VLamax [mmol/L/s]':        (0.45,  0.12,  True),
                'KH-Match [W/kg]':          (3.60,  0.55,  False),
                'Maximale Leistung [W/kg]': (15.5,  2.5,   False),
                'Fatmax [W/kg]':            (3.35,  0.50,  False),
                'ANS rel. [W/kg]':          (4.55,  0.55,  False),
            }
        },
        "weiblich": {
            "Hobby": {
                'ANS abs. [W]':             (150.0, 30.0,  False),
                'Fettanteil [%]':           (20.0,  3.0,   True),
                'VO2max [ml/min/kg]':       (42.0,  5.0,   False),
                'VLamax [mmol/L/s]':        (0.60,  0.14,  True),
                'KH-Match [W/kg]':          (1.85,  0.40,  False),
                'Maximale Leistung [W/kg]': (8.5,   2.0,   False),
                'Fatmax [W/kg]':            (1.65,  0.35,  False),
                'ANS rel. [W/kg]':          (2.30,  0.45,  False),
            },
            "Amateur": {
                'ANS abs. [W]':             (175.0, 40.0,  False),
                'Fettanteil [%]':           (17.0,  2.8,   True),
                'VO2max [ml/min/kg]':       (48.5,  5.0,   False),
                'VLamax [mmol/L/s]':        (0.50,  0.12,  True),
                'KH-Match [W/kg]':          (2.25,  0.45,  False),
                'Maximale Leistung [W/kg]': (9.2,   2.0,   False),
                'Fatmax [W/kg]':            (2.00,  0.45,  False),
                'ANS rel. [W/kg]':          (2.75,  0.55,  False),
            },
            "Profi": {
                'ANS abs. [W]':             (210.0, 40.0,  False),
                'Fettanteil [%]':           (13.5,  3.0,   True),
                'VO2max [ml/min/kg]':       (55.0,  5.0,   False),
                'VLamax [mmol/L/s]':        (0.42,  0.12,  True),
                'KH-Match [W/kg]':          (2.80,  0.50,  False),
                'Maximale Leistung [W/kg]': (11.5,  2.2,   False),
                'Fatmax [W/kg]':            (2.65,  0.50,  False),
                'ANS rel. [W/kg]':          (3.45,  0.50,  False),
            }
        }
    }

    category_params = ref_table[gender_key][cat_key]

    def normalize_ref(val, cat):
        mean, sd, inverted = category_params[cat]
        if sd <= 0:
            return 65.0
        z = (mean - val) / sd if inverted else (val - mean) / sd
        # Sanfte Sigmoid-/Linear-Normierung analog zur HYCYS-Referenzspinne (Mittelwert ~ 60 %, guter Bereich 60-85 %, kein Kollabieren < 45 %)
        raw_score = 60.0 + 13.5 * np.tanh(z / 1.65) + 4.5 * np.clip(z, -2.2, 2.2)
        return float(np.clip(raw_score, 44.0, 94.0))

    radar_cats = ['ANS abs. [W]', 'Fettanteil [%]', 'VO2max [ml/min/kg]', 'VLamax [mmol/L/s]',
                  'KH-Match [W/kg]', 'Maximale Leistung [W/kg]', 'Fatmax [W/kg]', 'ANS rel. [W/kg]']
    radar_labels = ['ANS abs. [W]', 'Fettanteil [%]', 'VO2max [ml/min/kg]', 'VLamax [mmol/L/s]',
                    'KH-Match [W/kg]', 'Maximale Leistung [W/kg]', 'Fatmax [W/kg]', 'ANS rel. [W/kg]']
    N      = len(radar_cats)
    angles = [n / float(N) * 2 * np.pi for n in range(N)]
    angles += angles[:1]

    fig, ax = plt.subplots(figsize=(5.0, 5.0), dpi=300)
    ax.set_aspect('equal')

    # Background rings exakt wie im HYCYS Original-Excel (100 / 80 / 60 / 35): Sehr gut > Gut > Zu verbessern > Schwach
    x_100 = [100 * np.cos(t) for t in angles]
    y_100 = [100 * np.sin(t) for t in angles]
    ax.fill(x_100, y_100, color='#cdb663', edgecolor='none', zorder=1)

    x_80 = [80 * np.cos(t) for t in angles]
    y_80 = [80 * np.sin(t) for t in angles]
    ax.fill(x_80, y_80, color='#2cb7b9', edgecolor='none', zorder=2)

    x_60 = [60 * np.cos(t) for t in angles]
    y_60 = [60 * np.sin(t) for t in angles]
    ax.fill(x_60, y_60, color='#ffffff', edgecolor='none', zorder=3)

    x_35 = [35 * np.cos(t) for t in angles]
    y_35 = [35 * np.sin(t) for t in angles]
    ax.fill(x_35, y_35, color='#595a59', edgecolor='none', zorder=4)

    # Draw spokes (radial lines) and inner ring helper boundaries so they are visible inside the dark-grey center
    for r_ring_line in [35, 60, 80]:
        ax.plot(
            [r_ring_line * np.cos(t) for t in angles],
            [r_ring_line * np.sin(t) for t in angles],
            color='#8c8d8c' if r_ring_line == 35 else '#b8b9b8',
            linestyle='-',
            linewidth=0.6,
            zorder=5,
        )
    for t in angles[:-1]:
        ax.plot([0, 35 * np.cos(t)], [0, 35 * np.sin(t)], color='#9e9f9e', linestyle='-', linewidth=0.8, zorder=5)
        ax.plot([35 * np.cos(t), 100 * np.cos(t)], [35 * np.sin(t), 100 * np.sin(t)], color='#595a59', linestyle='-', linewidth=0.8, zorder=5)

    # Draw outer octagon spine boundary
    ax.plot(x_100, y_100, color='#595a59', linewidth=1.2, zorder=8)

    # KH-Match in der HYCYS-Spinne wird immer mit 1,0 g KH / kg Körpergewicht / h berechnet!
    try:
        _kh_1g_target = 1.0 * float(weight)
        _df_khm_1g = df_sim[df_sim['ee_kh_g'] <= _kh_1g_target]
        if not _df_khm_1g.empty:
            _khm_1g_power = float(_df_khm_1g['power'].max())
        else:
            _khm_1g_power = float(carb_match_power)
    except Exception:
        _khm_1g_power = float(carb_match_power)

    ath_pmax_rel  = max_sprint_power / weight if weight > 0 else 0.0
    ath_match_rel = _khm_1g_power / weight if weight > 0 else 0.0

    athlete_scores = [
        normalize_ref(ans_power,     'ANS abs. [W]'),
        normalize_ref(body_fat_pct,  'Fettanteil [%]'),
        normalize_ref(rel_vo2max,    'VO2max [ml/min/kg]'),
        normalize_ref(vlamax,        'VLamax [mmol/L/s]'),
        normalize_ref(ath_match_rel, 'KH-Match [W/kg]'),
        normalize_ref(ath_pmax_rel,  'Maximale Leistung [W/kg]'),
        normalize_ref(fatmax_rel,    'Fatmax [W/kg]'),
        normalize_ref(ans_rel,       'ANS rel. [W/kg]'),
    ]
    athlete_scores += athlete_scores[:1]

    # Draw athlete line (1:1 wie in Inas HYCYS-PDF: markante schwarze Linie ohne graue Innenfüllung)
    x_ath = [r * np.cos(t) for r, t in zip(athlete_scores, angles)]
    y_ath = [r * np.sin(t) for r, t in zip(athlete_scores, angles)]
    ax.plot(x_ath, y_ath, color='#000000', linewidth=2.8, zorder=7)

    # Add text labels manually (Schriftgröße und Abstand exakt wie im Original-HYCYS-PDF)
    for i, label in enumerate(radar_labels):
        t = angles[i]
        r_label = 106
        x = r_label * np.cos(t)
        y = r_label * np.sin(t)
        
        ha = 'center'
        if np.cos(t) > 0.1:
            ha = 'left'
        elif np.cos(t) < -0.1:
            ha = 'right'
            
        va = 'center'
        if np.sin(t) > 0.1:
            va = 'bottom'
        elif np.sin(t) < -0.1:
            va = 'top'
            
        if abs(np.cos(t)) < 0.1:
            if np.sin(t) > 0:
                y += 1.5
            else:
                y -= 1.5
                
        ax.text(x, y, label, fontsize=7.6, fontweight='bold', color='#595a59', ha=ha, va=va)

    ax.axis('off')
    ax.set_xlim(-138, 138)
    ax.set_ylim(-128, 128)

    legend_handles = [
        Patch(facecolor='#cdb663', alpha=1.0,  label='Sehr gut'),
        Patch(facecolor='#2cb7b9', alpha=1.0,  label='Gut'),
        Patch(facecolor='#ffffff', edgecolor='#cccccc', alpha=1.0,  label='Zu verbessern'),
        Patch(facecolor='#595a59', alpha=1.0,  label='Schwach'),
        Line2D([0], [0], color='#000000', lw=2.5, label='Daten Sportler'),
        Line2D([0], [0], color='#595a59', lw=1.8, ls='--', label='Coaching Potential'),
    ]
    legend = ax.legend(
        handles=legend_handles,
        loc='lower center',
        bbox_to_anchor=(0.5, -0.06),
        ncol=6,
        fontsize=6.4,
        frameon=False,
        handlelength=1.1,
        handletextpad=0.4,
        columnspacing=1.0,
    )
    for text in legend.get_texts():
        text.set_color('#1a1a1a')
        text.set_weight('bold')
    fig.patch.set_facecolor('white')
    plt.tight_layout()
    with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp:
        fig.savefig(tmp.name, format='png', bbox_inches='tight', facecolor='white')
        plt.close(fig)
        
        # Pad the image proportionally to avoid clipping wide labels while keeping the octagon large and above Trainingsbereiche
        from PIL import Image
        img = Image.open(tmp.name)
        w_img, h_img = img.size
        _w_spider_pt = 300.0
        _h_spider_pt = _w_spider_pt * (h_img / float(w_img))
        pdf.image(tmp.name, x=140.0, y=144.0, w=_w_spider_pt, h=_h_spider_pt)
        temp_files.append(tmp.name)

    # Training zones table
    draw_rect(57.6, 414.9, 155.3, 10.44, fill_color=c_teal)
    draw_text_centered(135.3, 416.1, "Trainingsbereiche", size=8.2, font_style='B', color=c_white)

    # Table header
    draw_text(99.1,  435.7, "Trainingsbereich", size=8.2, font_style='B', color=c_dark_grey)
    draw_text(256.9, 435.7, "Leistung",         size=8.2, font_style='B', color=c_dark_grey)
    draw_text(382.8, 435.7, "Herzfrequenz",     size=8.2, font_style='B', color=c_dark_grey)
    draw_text(481.9, 435.7, "Tretfreq.",        size=8.2, font_style='B', color=c_dark_grey)
    draw_text(263.8, 446.1, "[Watt]",           size=7.6, color=c_dark_grey)
    draw_text(399.1, 446.1, "[bpm]",            size=7.6, color=c_dark_grey)
    draw_text(486.3, 446.1, "[U/min]",          size=7.6, color=c_dark_grey)
    draw_text(222.4, 456.6, "min",  size=7.6, color=c_dark_grey)
    draw_text(266.5, 456.6, "max",  size=7.6, color=c_dark_grey)
    draw_text(312.3, 456.6, "Ziel", size=7.6, color=c_dark_grey)
    draw_text(357.1, 456.6, "min",  size=7.6, color=c_dark_grey)
    draw_text(401.1, 456.6, "max",  size=7.6, color=c_dark_grey)
    draw_text(447.0, 456.6, "Ziel", size=7.6, color=c_dark_grey)

    def hr_at_p(p):
        return int(round(min(hfmax_val, max(40, p * slope_val + intercept_val))))

    # Zone boundaries (1:1 wie im Original HYCYS Report_Cycling Excel: ungerundete Prozentwerte von ANS bzw. MAP)
    kb_max_val = int(round(ans_power * 0.50))
    kb_zl_val  = int(round((ans_power * 0.50 * 0.90) / 10.0) * 10)
    g1_min_val = kb_max_val
    g1_max_val = int(round(ans_power * 0.75))
    try:
        g1_sub = df_sim[df_sim['ee_kh_g'] <= 1.0 * weight]
        if not g1_sub.empty:
            g1_zl_val = int(round(float(g1_sub['power'].max()) / 10.0) * 10)
        else:
            g1_i = (df_sim['ee_kh_g'] - 1.0 * weight).abs().idxmin()
            g1_zl_val = int(round(float(df_sim.loc[g1_i, 'power']) / 10.0) * 10)
    except Exception:
        g1_zl_val = int(round((ans_power * 0.72) / 10.0) * 10)
    g2_min_val = g1_max_val
    g2_max_val = int(round(ans_power * 0.90))
    g2_zl_val  = int(round((ans_power * 0.86) / 10.0) * 10)
    eb_min_val = g2_max_val
    eb_max_val = int(round(ans_power * 1.05))
    eb_zl_val  = int(round(ans_power / 10.0) * 10)
    sb_min_val = eb_max_val
    sb_max_val = int(round(map_val * 1.10))
    sb_zl_val  = int(round((map_val * 1.075) / 10.0) * 10)
    k3_min_val = int(round(ans_power * 0.84))
    k3_max_val = int(round(ans_power * 0.98))
    k3_zl_val  = int(round((ans_power * 0.92) / 10.0) * 10)

    # Calculate heart rates at boundaries and target power (1:1 wie LOOKUP in Calculation!$E$9:$AQ$609)
    hr_kb_max = hr_at_p(ans_power * 0.50)
    hr_kb_zl  = hr_at_p(kb_zl_val)
    hr_g1_min = hr_kb_max
    hr_g1_max = hr_at_p(ans_power * 0.75)
    hr_g1_zl  = hr_at_p(g1_zl_val)
    hr_g2_min = hr_g1_max
    hr_g2_max = hr_at_p(ans_power * 0.90)
    hr_g2_zl  = hr_at_p(g2_zl_val)
    hr_eb_min = hr_g2_max
    hr_eb_max = hr_at_p(ans_power * 1.05)
    hr_eb_zl  = hr_at_p(eb_zl_val)
    hr_sb_min = hr_eb_max
    hr_sb_max = hr_at_p(map_val * 1.10)
    hr_sb_zl  = hr_at_p(sb_zl_val)
    hr_k3_min = hr_at_p(ans_power * 0.84)
    hr_k3_max = hr_at_p(ans_power * 0.98)
    hr_k3_zl  = hr_at_p(k3_zl_val)

    def rnd(v): return int(round(v / 10.0) * 10)

    zones = [
        ("Kompensationsbereich (KB)", "-",        kb_max_val, kb_zl_val,
         "-",              hr_kb_max,        hr_kb_zl,             "80-100"),
        ("Grundlagenausdauer 1 (G1)", g1_min_val, g1_max_val, g1_zl_val,
         hr_g1_min,        hr_g1_max,        hr_g1_zl,             "85-110"),
        ("Grundlagenausdauer 2 (G2)", g2_min_val, g2_max_val, g2_zl_val,
         hr_g2_min,        hr_g2_max,        hr_g2_zl,             "85-110"),
        ("Entwicklungsbereich (EB)",  eb_min_val, eb_max_val, eb_zl_val,
         hr_eb_min,        hr_eb_max,        hr_eb_zl,             "85-110"),
        ("Spitzenbereich (SB)",       sb_min_val, sb_max_val, sb_zl_val,
         hr_sb_min,        hr_sb_max,        hr_sb_zl,             "95-120"),
        ("Kraftausdauer (K3)",        k3_min_val, k3_max_val, k3_zl_val,
         hr_k3_min,        hr_k3_max,        hr_k3_zl,             "40-60"),
    ]

    y_row = 466.2
    for name, p_min, p_max, p_zl, hr_mn, hr_mx, hr_zl, tf in zones:
        draw_text(58.7,  y_row, name,       size=7.6, font_style='B', color=c_dark_grey)
        if str(p_min) == "-":
            draw_text_centered(222.4 + 5.7, y_row, "-", size=7.6, color=c_dark_grey)
        else:
            draw_text(222.4, y_row, str(p_min), size=7.6, color=c_dark_grey)
        draw_text(266.5, y_row, str(p_max), size=7.6, color=c_dark_grey)
        draw_text(312.3, y_row, str(p_zl),  size=7.6, color=c_dark_grey)
        if str(hr_mn) == "-":
            draw_text_centered(357.1 + 5.7, y_row, "-", size=7.6, color=c_dark_grey)
        else:
            draw_text(357.1, y_row, str(hr_mn), size=7.6, color=c_dark_grey)
        draw_text(401.1, y_row, str(hr_mx), size=7.6, color=c_dark_grey)
        draw_text(447.0, y_row, str(hr_zl), size=7.6, color=c_dark_grey)
        draw_text(486.7, y_row, tf,          size=7.6, color=c_dark_grey)
        y_row += 9.9

    # Draw vertical separator lines for main table
    for x in [205.97, 340.63, 475.30, 520.20]:
        draw_rect(x, 455.69, 0.6, 69.74, fill_color=c_dark_grey)
    # Draw bottom border for main table
    draw_rect(206.57, 524.83, 314.23, 0.6, fill_color=c_dark_grey)

    # IE / LC sub-table
    draw_text(279.3, 535.1, "Leistung",  size=8.2, font_style='B', color=c_dark_grey)
    draw_text(481.9, 535.1, "Tretfreq.", size=8.2, font_style='B', color=c_dark_grey)
    draw_text(286.2, 545.9, "[Watt]",   size=7.6, color=c_dark_grey)
    draw_text(486.3, 545.9, "[U/min]",  size=7.6, color=c_dark_grey)
    draw_text(259.4, 555.8, "Intervall", size=7.6, color=c_dark_grey)
    draw_text(307.8, 555.8, "Pause",    size=7.6, color=c_dark_grey)

    ie_int   = rnd(map_val * 1.05)
    ie_pause = rnd(ans_power * 0.5)
    draw_text(58.8,  565.6, "Intermitted exercise (IE)", size=8.2, font_style='B', color=c_dark_grey)
    draw_text(267.6, 565.6, str(ie_int),   size=7.6, color=c_dark_grey)
    draw_text(312.5, 565.6, str(ie_pause), size=7.6, color=c_dark_grey)
    draw_text(486.7, 565.6, "85-110",      size=7.6, color=c_dark_grey)

    try:
        lc_sim = df_sim.copy()
        lc_above = lc_sim[lc_sim['vlass'] > lc_sim['vlaoxmax']]
        if not lc_above.empty:
            lc_above = lc_above.copy()
            lc_above['d1'] = (lc_above['vlass'] - lc_above['vlaoxmax'] - 1.0).abs()
            lc_int = int(round(float(lc_above.loc[lc_above['d1'].idxmin(), 'power'])))
        else:
            lc_int = 0
    except Exception:
        lc_int = 0
    lc_pause = int(round(fatmax_power))
    draw_text(58.8,  575.9, "Laktat Auf-/Abbau (LC)", size=8.2, font_style='B', color=c_dark_grey)
    draw_text(267.6, 575.9, str(lc_int),   size=7.6, color=c_dark_grey)
    draw_text(312.5, 575.9, str(lc_pause), size=7.6, color=c_dark_grey)
    draw_text(486.7, 575.9, "85-110",      size=7.6, color=c_dark_grey)

    # Draw vertical separator lines for sub-table
    for x in [250.85, 340.63, 475.30, 520.20]:
        draw_rect(x, 564.91, 0.6, 21.0, fill_color=c_dark_grey)
    # Draw bottom borders for sub-table
    draw_rect(251.45, 585.31, 89.78, 0.6, fill_color=c_dark_grey)
    draw_rect(475.90, 585.31, 44.90, 0.6, fill_color=c_dark_grey)

    # =========================================================
    # SEITE 5 – SPRINT & ANTRITT
    # =========================================================
    pdf.add_page()
    draw_rect(50.9, 54.5, 479.1, 663.6, fill_color=c_white)
    page_logo(5)

    # Large logo centred (as in reference page 5)
    if os.path.exists('pdf_assets/page1_img3.png'):
        pdf.image('pdf_assets/page1_img3.png', x=112.3, y=155.0, w=362.5, h=112.0)

    # Sprint & Antritt header bar
    draw_rect(57.6, 414.9, 458.1, 10.44, fill_color=c_teal)
    draw_text_centered(286.7, 416.1, "Sprint & Antritt", size=8.2, font_style='B', color=c_white)

    # Leistungswerte section
    draw_rect(57.6, 452.0, 143.6, 10.44, fill_color=c_gold)
    draw_text_centered(129.4, 453.2, "Leistungswerte", size=8.2, font_style='B', color=c_white)

    avg_p_str     = f"{int(round(avg_sprint_power))} W"
    max_p_str     = f"{int(round(max_sprint_power))} W"
    max_p_rel_str = (f"{max_sprint_power / weight:.1f} W/kg".replace('.', ',')
                     if weight > 0 else "0,0 W/kg")

    draw_text(58.7, 466.2, "durchschn. Leistung", size=7.6, font_style='B', color=c_black)
    draw_text(172.8, 466.7, avg_p_str,             size=7.6, color=c_black)
    draw_text(58.7, 476.1, "max. Leistung abs.",  size=7.6, font_style='B', color=c_black)
    draw_text(170.7, 476.6, max_p_str,             size=7.6, color=c_black)
    draw_text(58.7, 485.9, "max. Leistung rel.",  size=7.6, font_style='B', color=c_black)
    draw_text(166.6, 486.4, max_p_rel_str,         size=7.6, color=c_black)

    # Beschleunigung & Kraft section
    draw_rect(57.6, 531.2, 143.6, 10.44, fill_color=c_gold)
    draw_text_centered(129.4, 532.4, "Beschleunigung & Kraft", size=8.2, font_style='B', color=c_white)

    max_f_str     = f"{int(round(max_force))} N"
    f_at_pmax_str = f"{int(round(force_at_pmax))} N"
    f_pct_str     = f"{int(round(force_pct_max))} %"

    draw_text(58.7, 545.5, "max. Kraft",              size=7.6, font_style='B', color=c_black)
    draw_text(173.8, 545.9, max_f_str,                 size=7.6, color=c_black)
    draw_text(58.7, 555.3, "Kraft bei max. Leistung", size=7.6, font_style='B', color=c_black)
    draw_text(173.8, 555.8, f_at_pmax_str,             size=7.6, color=c_black)
    draw_text(58.7, 565.4, "% der max. Kraft",        size=7.6, font_style='B', color=c_black)
    draw_text(176.3, 565.9, f_pct_str,                 size=7.6, color=c_black)

    # Sprint Power vs Time chart
    fig, ax = plt.subplots(figsize=(4.4444, 2.5), dpi=300)
    ax.plot(sprint_times, sprint_powers, color='#595a59', linewidth=2.5)
    try:
        if t_alak > 0 and t_bel > t_alak:
            ax.axvspan(0.0,    t_alak, color=(0.0, 0.631, 0.878), alpha=0.18)
            ax.axvspan(t_alak, t_bel,  color=(0.804, 0.714, 0.388), alpha=0.18)
    except Exception:
        pass

    ax.set_ylabel('Leistung [W]', color='#595a59', fontsize=6, fontweight='bold')
    ax.tick_params(axis='y', labelcolor='#595a59', labelsize=6, width=1.2)
    ax.set_xlabel('Zeit [s]', color='#595a59', fontsize=6, fontweight='bold')
    ax.tick_params(axis='x', labelcolor='#595a59', labelsize=6, width=1.2)
    ax.grid(True, which='both', linestyle='-', linewidth=1.0, color='#d8d8d8')
    for spine in ax.spines.values():
        spine.set_color('#595a59')
        spine.set_linewidth(1.2)
    fig.patch.set_facecolor('white')
    plt.tight_layout()
    with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp:
        fig.savefig(tmp.name, format='png', bbox_inches='tight', facecolor='white')
        plt.close(fig)
        # Chart placed fully to the right of the value boxes, below the header bar
        pdf.image(tmp.name, x=220.0, y=435.0, w=288.0, h=162.0)
        temp_files.append(tmp.name)

    # Output bytes
    pdf_output = pdf.output(dest='S')
    for f in temp_files:
        try:
            os.unlink(f)
        except Exception:
            pass
    return pdf_output.encode('latin-1', 'replace')
