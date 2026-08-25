import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import os
import random
import re

from datetime import datetime
from io import BytesIO

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


# ============================================================
# CONFIGURATION
# ============================================================

EXCEL_FILE = r"quotation_responses.xlsx"

st.set_page_config(
    page_title="BMW Vehicle Quotation Request",
    page_icon="🚘",
    layout="centered"
)


# ============================================================
# SESSION STATE
# ============================================================

if "scratch_discount" not in st.session_state:

    st.session_state["scratch_discount"] = random.choice(
        list(
            range(
                25000,
                100001,
                5000
            )
        )
    )


if "pdf_file" not in st.session_state:

    st.session_state["pdf_file"] = None


if "pdf_name" not in st.session_state:

    st.session_state["pdf_name"] = None


if "purchase_continue" not in st.session_state:

    st.session_state["purchase_continue"] = False


# ============================================================
# INDIAN RUPEE FORMATTER
# ============================================================

def format_inr(amount):

    try:
        amount = int(float(amount))

    except Exception:
        return "₹0"

    amount_str = str(abs(amount))

    if len(amount_str) <= 3:

        formatted = amount_str

    else:

        last_three = amount_str[-3:]

        remaining = amount_str[:-3]

        parts = []

        while len(remaining) > 2:

            parts.insert(
                0,
                remaining[-2:]
            )

            remaining = remaining[:-2]

        if remaining:

            parts.insert(
                0,
                remaining
            )

        formatted = (
            ",".join(parts)
            +
            ","
            +
            last_three
        )

    if amount < 0:

        return "-₹" + formatted

    return "₹" + formatted


# ============================================================
# PDF FONT
# ============================================================

def register_pdf_font():

    possible_fonts = [

        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",

        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",

        "C:/Windows/Fonts/DejaVuSans.ttf",

        "C:/Windows/Fonts/arial.ttf",

        "C:/Windows/Fonts/ARIAL.TTF"

    ]

    for font_path in possible_fonts:

        if os.path.exists(font_path):

            try:

                pdfmetrics.registerFont(
                    TTFont(
                        "LuxuryUnicode",
                        font_path
                    )
                )

                return "LuxuryUnicode"

            except Exception:
                pass

    try:

        import matplotlib.font_manager as fm

        font_path = fm.findfont(
            "DejaVu Sans"
        )

        if os.path.exists(font_path):

            pdfmetrics.registerFont(
                TTFont(
                    "LuxuryUnicode",
                    font_path
                )
            )

            return "LuxuryUnicode"

    except Exception:
        pass

    return "Times-Roman"


PDF_FONT = register_pdf_font()


# ============================================================
# PREMIUM LUXURY UI
# ============================================================

st.markdown(
    """
<style>

/* ============================================================
   GOOGLE LUXURY FONTS
   ============================================================ */

@import url(
    'https://fonts.googleapis.com/css2?family=Cormorant+Garamond:wght@400;500;600;700&family=Inter:wght@400;500;600;700&display=swap'
);


/* ============================================================
   GLOBAL
   ============================================================ */

.stApp {

    background:
        linear-gradient(
            135deg,
            #EEE6DA 0%,
            #E7DCCB 48%,
            #EFE8DE 100%
        );

    color:
        #29241F;
}


.block-container {

    max-width:
        1020px !important;

    padding-top:
        1.4rem !important;

    padding-bottom:
        3rem !important;
}


#MainMenu {

    visibility:
        hidden;
}


footer {

    visibility:
        hidden !important;
}


header {

    visibility:
        hidden !important;
}


/* ============================================================
   GLOBAL FONT
   ============================================================ */

html,
body,
.stApp,
[data-testid="stAppViewContainer"] {

    font-family:
        "Inter",
        Arial,
        sans-serif;
}


/* ============================================================
   LUXURY HEADER
   ============================================================ */

.luxury-header {

    position:
        relative;

    text-align:
        center;

    margin-top:
        8px;

    margin-bottom:
        42px;

    padding:
        38px 28px 35px 28px;

    background:
        linear-gradient(
            145deg,
            #181614 0%,
            #28231F 42%,
            #171513 100%
        );

    border:
        1px solid #8D7044;

    border-radius:
        8px;

    box-shadow:
        0 18px 45px
        rgba(35,28,21,0.28);

    overflow:
        hidden;
}


.luxury-header::before {

    content:
        "";

    position:
        absolute;

    top:
        8px;

    left:
        8px;

    right:
        8px;

    bottom:
        8px;

    border:
        1px solid
        rgba(210,174,103,0.25);

    border-radius:
        5px;

    pointer-events:
        none;
}


.header-top-line {

    position:
        absolute;

    top:
        0;

    left:
        0;

    right:
        0;

    height:
        3px;

    background:
        linear-gradient(
            90deg,
            transparent,
            #A77E3F,
            #F0D28B,
            #A77E3F,
            transparent
        );
}


/* ============================================================
   BMW WORDMARK
   ============================================================ */

.bmw-brand {

    color:
        #F8F1E7;

    font-family:
        "Cormorant Garamond",
        Georgia,
        serif;

    font-size:
        52px;

    font-weight:
        600;

    letter-spacing:
        14px;

    line-height:
        1;

    margin-left:
        14px;

    text-shadow:
        0 3px 12px
        rgba(0,0,0,0.55);
}


.bmw-subtitle {

    color:
        #D2B06B;

    font-family:
        "Inter",
        Arial,
        sans-serif;

    font-size:
        10px;

    font-weight:
        600;

    letter-spacing:
        5px;

    margin-top:
        14px;
}


/* ============================================================
   HEADER DIVIDER
   ============================================================ */

.header-divider {

    display:
        flex;

    align-items:
        center;

    justify-content:
        center;

    gap:
        14px;

    max-width:
        330px;

    margin:
        20px auto 17px auto;
}


.header-divider span {

    height:
        1px;

    flex:
        1;

    background:
        linear-gradient(
            90deg,
            transparent,
            #9C7A45
        );
}


.header-divider span:last-child {

    background:
        linear-gradient(
            90deg,
            #9C7A45,
            transparent
        );
}


.divider-diamond {

    color:
        #D8B76C;

    font-size:
        9px;
}


.bmw-tagline {

    color:
        #F0E6D8;

    font-family:
        "Cormorant Garamond",
        Georgia,
        serif;

    font-size:
        19px;

    font-weight:
        500;

    font-style:
        italic;

    letter-spacing:
        2.5px;
}


.bmw-description {

    color:
        #BEB1A0;

    font-family:
        "Inter",
        Arial,
        sans-serif;

    font-size:
        9px;

    letter-spacing:
        1.8px;

    margin-top:
        10px;
}


/* ============================================================
   FORM LABELS
   ============================================================ */

.stTextInput label,
.stSelectbox label,
.stTextArea label,
.stDateInput label,
.stNumberInput label,
.stRadio > label,
.stMultiSelect label,
.stCheckbox label {

    color:
        #51473C !important;

    font-family:
        "Inter",
        Arial,
        sans-serif !important;

    font-size:
        11px !important;

    font-weight:
        600 !important;

    letter-spacing:
        0.7px !important;
}


/* ============================================================
   TEXT INPUTS
   ============================================================ */

.stTextInput > div > div {

    background:
        #FFFDF9 !important;

    border:
        1px solid #D6C9B8 !important;

    border-radius:
        4px !important;

    box-shadow:
        0 2px 8px
        rgba(44,35,27,0.035) !important;

    transition:
        all 0.25s ease !important;
}


.stTextInput > div > div:hover {

    border-color:
        #B18A4A !important;

    box-shadow:
        0 4px 12px
        rgba(177,138,74,0.09) !important;
}


.stTextInput > div > div:focus-within {

    border-color:
        #A77E3F !important;

    box-shadow:
        0 0 0 2px
        rgba(167,126,63,0.12) !important;
}


.stTextInput input {

    background:
        #FFFDF9 !important;

    color:
        #29241F !important;

    -webkit-text-fill-color:
        #29241F !important;

    font-family:
        "Cormorant Garamond",
        Georgia,
        serif !important;

    font-size:
        16px !important;

    font-weight:
        500 !important;

    letter-spacing:
        0.2px !important;
}


.stTextInput input::placeholder {

    color:
        #9A8E80 !important;

    -webkit-text-fill-color:
        #9A8E80 !important;

    opacity:
        1 !important;
}


/* ============================================================
   NUMBER INPUT
   ============================================================ */

.stNumberInput > div > div {

    background:
        #FFFDF9 !important;

    border:
        1px solid #D6C9B8 !important;

    border-radius:
        4px !important;

    box-shadow:
        0 2px 8px
        rgba(44,35,27,0.035) !important;
}


.stNumberInput > div > div:focus-within {

    border-color:
        #A77E3F !important;

    box-shadow:
        0 0 0 2px
        rgba(167,126,63,0.12) !important;
}


.stNumberInput input {

    background:
        #FFFDF9 !important;

    color:
        #29241F !important;

    -webkit-text-fill-color:
        #29241F !important;

    font-family:
        "Cormorant Garamond",
        Georgia,
        serif !important;

    font-size:
        16px !important;

    font-weight:
        600 !important;
}


/* Number buttons */

.stNumberInput button {

    color:
        #9C7A45 !important;

    background:
        transparent !important;

    border:
        none !important;
}


/* ============================================================
   ENQUIRY DATE - IMPORTANT FIX
   ============================================================ */

.stDateInput [data-baseweb="input"] {

    background:
        #FFFDF9 !important;

    border:
        1px solid #D6C9B8 !important;

    border-radius:
        4px !important;

    box-shadow:
        0 2px 8px
        rgba(44,35,27,0.035) !important;

    transition:
        all 0.25s ease !important;
}


.stDateInput [data-baseweb="input"]:hover {

    border-color:
        #B18A4A !important;

    box-shadow:
        0 4px 12px
        rgba(177,138,74,0.09) !important;
}


.stDateInput [data-baseweb="input"]:focus-within {

    border-color:
        #A77E3F !important;

    box-shadow:
        0 0 0 2px
        rgba(167,126,63,0.12) !important;
}


.stDateInput input {

    background:
        #FFFDF9 !important;

    color:
        #29241F !important;

    -webkit-text-fill-color:
        #29241F !important;

    opacity:
        1 !important;

    font-family:
        "Cormorant Garamond",
        Georgia,
        serif !important;

    font-size:
        16px !important;

    font-weight:
        600 !important;
}


.stDateInput input::placeholder {

    color:
        #8B7F71 !important;

    -webkit-text-fill-color:
        #8B7F71 !important;

    opacity:
        1 !important;
}


/* Calendar icon */

.stDateInput svg {

    color:
        #A77E3F !important;

    fill:
        #A77E3F !important;

    stroke:
        #A77E3F !important;
}


/* ============================================================
   SELECT BOX
   ============================================================ */

.stSelectbox div[data-baseweb="select"] {

    background:
        #FFFDF9 !important;

    border:
        1px solid #D6C9B8 !important;

    border-radius:
        4px !important;

    min-height:
        44px !important;

    box-shadow:
        0 2px 8px
        rgba(44,35,27,0.035) !important;

    transition:
        all 0.25s ease !important;
}


.stSelectbox div[data-baseweb="select"]:hover {

    border-color:
        #B18A4A !important;
}


.stSelectbox div[data-baseweb="select"]:focus-within {

    border-color:
        #A77E3F !important;

    box-shadow:
        0 0 0 2px
        rgba(167,126,63,0.12) !important;
}


.stSelectbox div[data-baseweb="select"] * {

    color:
        #29241F !important;

    font-family:
        "Cormorant Garamond",
        Georgia,
        serif !important;

    font-size:
        15px !important;

    font-weight:
        500 !important;
}


/* Select arrow */

.stSelectbox svg {

    color:
        #A77E3F !important;
}


/* ============================================================
   MULTISELECT
   ============================================================ */

.stMultiSelect div[data-baseweb="select"] {

    background:
        #FFFDF9 !important;

    border:
        1px solid #D6C9B8 !important;

    border-radius:
        4px !important;

    min-height:
        46px !important;

    box-shadow:
        0 2px 8px
        rgba(44,35,27,0.035) !important;
}


.stMultiSelect div[data-baseweb="select"]:focus-within {

    border-color:
        #A77E3F !important;

    box-shadow:
        0 0 0 2px
        rgba(167,126,63,0.12) !important;
}


.stMultiSelect div[data-baseweb="select"] * {

    font-family:
        "Inter",
        Arial,
        sans-serif !important;

    color:
        #29241F !important;
}


/* Selected tags */

.stMultiSelect span[data-baseweb="tag"] {

    background:
        #E7D6B5 !important;

    border:
        1px solid #B18A4A !important;

    color:
        #49371F !important;

    border-radius:
        3px !important;
}


/* ============================================================
   TEXT AREA
   ============================================================ */

.stTextArea > div > div {

    background:
        #FFFDF9 !important;

    border:
        1px solid #D6C9B8 !important;

    border-radius:
        4px !important;

    box-shadow:
        0 2px 8px
        rgba(44,35,27,0.035) !important;
}


.stTextArea > div > div:focus-within {

    border-color:
        #A77E3F !important;

    box-shadow:
        0 0 0 2px
        rgba(167,126,63,0.12) !important;
}


.stTextArea textarea {

    background:
        #FFFDF9 !important;

    color:
        #29241F !important;

    -webkit-text-fill-color:
        #29241F !important;

    font-family:
        "Cormorant Garamond",
        Georgia,
        serif !important;

    font-size:
        16px !important;

    line-height:
        1.5 !important;
}


.stTextArea textarea::placeholder {

    color:
        #9A8E80 !important;

    -webkit-text-fill-color:
        #9A8E80 !important;
}


/* ============================================================
   SECTION TITLES
   ============================================================ */

.section-title {

    position:
        relative;

    background:
        linear-gradient(
            100deg,
            #26211D,
            #383029,
            #26211D
        );

    color:
        #F6EBDD;

    padding:
        14px 20px 13px 22px;

    font-family:
        "Cormorant Garamond",
        Georgia,
        serif;

    font-size:
        17px;

    font-weight:
        600;

    letter-spacing:
        3px;

    border:
        1px solid
        rgba(190,153,85,0.28);

    border-left:
        4px solid #C7A15D;

    margin-top:
        32px;

    margin-bottom:
        19px;

    border-radius:
        4px;

    box-shadow:
        0 6px 16px
        rgba(0,0,0,0.11);
}


.section-title::after {

    content:
        "◆";

    position:
        absolute;

    right:
        17px;

    top:
        50%;

    transform:
        translateY(-50%);

    color:
        #C8A45D;

    font-size:
        7px;
}


/* ============================================================
   RADIO / EXCHANGE OPTIONS
   ============================================================ */

.stRadio > div {

    display:
        flex !important;

    flex-direction:
        row !important;

    gap:
        12px !important;

    align-items:
        stretch !important;

    margin-top:
        5px !important;
}


/* Question */

.stRadio > label {

    color:
        #51473C !important;

    font-family:
        "Inter",
        Arial,
        sans-serif !important;

    font-size:
        11px !important;

    font-weight:
        600 !important;

    letter-spacing:
        0.7px !important;
}


/* Individual option */

.stRadio > div > label {

    position:
        relative !important;

    display:
        flex !important;

    align-items:
        center !important;

    justify-content:
        center !important;

    min-width:
        115px !important;

    min-height:
        46px !important;

    padding:
        8px 20px !important;

    background:
        #FFFDF9 !important;

    border:
        1px solid #D6C9B8 !important;

    border-radius:
        5px !important;

    color:
        #302A24 !important;

    font-family:
        "Inter",
        Arial,
        sans-serif !important;

    font-size:
        12px !important;

    font-weight:
        600 !important;

    letter-spacing:
        0.5px !important;

    cursor:
        pointer !important;

    box-shadow:
        0 2px 8px
        rgba(44,35,27,0.04) !important;

    transition:
        all 0.25s ease !important;
}


/* Option text */

.stRadio > div > label p {

    color:
        #302A24 !important;

    font-family:
        "Inter",
        Arial,
        sans-serif !important;

    font-size:
        12px !important;

    font-weight:
        600 !important;

    margin:
        0 !important;
}


/* Hover */

.stRadio > div > label:hover {

    background:
        #F5EDDF !important;

    border-color:
        #B18A4A !important;

    box-shadow:
        0 5px 14px
        rgba(177,138,74,0.13) !important;

    transform:
        translateY(-1px);
}


/* Hide default radio appearance visually */

.stRadio input {

    accent-color:
        #A77E3F !important;
}


/* Selected option */

.stRadio > div > label:has(input:checked) {

    background:
        linear-gradient(
            135deg,
            #F3E6CF,
            #E6D0A8
        ) !important;

    border:
        1px solid #A77E3F !important;

    box-shadow:
        0 6px 17px
        rgba(167,126,63,0.18) !important;
}


/* Selected option text */

.stRadio > div > label:has(input:checked) p {

    color:
        #513D20 !important;

    font-weight:
        700 !important;
}


/* ============================================================
   CHECKBOX
   ============================================================ */

.stCheckbox {

    margin-top:
        10px !important;
}


.stCheckbox label {

    color:
        #433A31 !important;

    font-family:
        "Inter",
        Arial,
        sans-serif !important;

    font-size:
        11px !important;

    font-weight:
        500 !important;
}


.stCheckbox input {

    accent-color:
        #A77E3F !important;
}


/* ============================================================
   BUTTONS
   ============================================================ */

.stButton > button,
.stFormSubmitButton button,
.stDownloadButton button {

    min-height:
        50px !important;

    border-radius:
        4px !important;

    font-family:
        "Inter",
        Arial,
        sans-serif !important;

    font-size:
        11px !important;

    font-weight:
        700 !important;

    letter-spacing:
        2px !important;

    transition:
        all 0.25s ease !important;
}


/* Standard button */

.stButton > button {

    background:
        linear-gradient(
            135deg,
            #A77E3F,
            #C29B57
        ) !important;

    color:
        #FFFFFF !important;

    border:
        1px solid #967038 !important;

    box-shadow:
        0 5px 14px
        rgba(130,96,43,0.18) !important;
}


.stButton > button:hover {

    background:
        linear-gradient(
            135deg,
            #29241F,
            #40362C
        ) !important;

    border-color:
        #B99557 !important;

    transform:
        translateY(-1px) !important;
}


/* ============================================================
   FORM SUBMIT BUTTON
   ============================================================ */

.stFormSubmitButton button {

    min-height:
        55px !important;

    background:
        linear-gradient(
            135deg,
            #211D19,
            #3B332B
        ) !important;

    color:
        #F5EBDD !important;

    border:
        1px solid #A77E3F !important;

    box-shadow:
        0 7px 18px
        rgba(31,25,20,0.18) !important;
}


.stFormSubmitButton button:hover {

    background:
        linear-gradient(
            135deg,
            #A77E3F,
            #C09A59
        ) !important;

    color:
        #FFFFFF !important;

    border-color:
        #D5B776 !important;

    transform:
        translateY(-1px) !important;
}


/* ============================================================
   DOWNLOAD BUTTON
   ============================================================ */

.stDownloadButton button {

    min-height:
        55px !important;

    background:
        linear-gradient(
            135deg,
            #A77E3F,
            #C09A59
        ) !important;

    color:
        #FFFFFF !important;

    border:
        1px solid #967038 !important;

    box-shadow:
        0 7px 18px
        rgba(130,96,43,0.20) !important;
}


.stDownloadButton button:hover {

    background:
        linear-gradient(
            135deg,
            #29241F,
            #40362C
        ) !important;

    border-color:
        #B99557 !important;

    transform:
        translateY(-1px) !important;
}


/* ============================================================
   CONTINUE STATUS
   ============================================================ */

.continue-status {

    text-align:
        center;

    color:
        #735C39;

    font-family:
        "Inter",
        Arial,
        sans-serif;

    font-size:
        10px;

    font-weight:
        600;

    letter-spacing:
        1.8px;

    margin-top:
        10px;

    margin-bottom:
        6px;
}


/* ============================================================
   DISCOUNT INFORMATION
   ============================================================ */

.discount-info {

    text-align:
        center;

    color:
        #765E3B;

    font-family:
        "Inter",
        Arial,
        sans-serif;

    font-size:
        10px;

    letter-spacing:
        1.4px;

    margin-top:
        12px;

    margin-bottom:
        10px;
}


.discount-info b {

    color:
        #8A642B;

    font-weight:
        700;
}


/* ============================================================
   ALERTS
   ============================================================ */

div[data-testid="stAlert"] {

    border-radius:
        5px !important;

    font-family:
        "Inter",
        Arial,
        sans-serif !important;

    font-size:
        12px !important;
}


/* ============================================================
   SUCCESS MESSAGE
   ============================================================ */

div[data-testid="stAlert"][data-baseweb="notification"] {

    border-left:
        3px solid #A77E3F !important;
}


/* ============================================================
   QUOTATION DOWNLOAD HEADING
   ============================================================ */

.stMarkdown h3 {

    color:
        #302A24 !important;

    font-family:
        "Cormorant Garamond",
        Georgia,
        serif !important;

    font-size:
        24px !important;

    font-weight:
        600 !important;

    letter-spacing:
        1.5px !important;
}


/* ============================================================
   LUXURY FOOTER
   ============================================================ */

.luxury-footer {

    text-align:
        center;

    margin-top:
        65px;

    margin-bottom:
        15px;

    padding:
        30px 20px 22px 20px;

    border-top:
        1px solid
        rgba(139,112,70,0.16);
}


.footer-divider {

    display:
        flex;

    align-items:
        center;

    justify-content:
        center;

    gap:
        13px;

    max-width:
        290px;

    margin:
        0 auto 22px auto;
}


.footer-divider span {

    height:
        1px;

    flex:
        1;

    background:
        linear-gradient(
            90deg,
            transparent,
            #A98A59
        );
}


.footer-divider span:last-child {

    background:
        linear-gradient(
            90deg,
            #A98A59,
            transparent
        );
}


.footer-diamond {

    color:
        #9C7A45;

    font-size:
        8px;
}


.footer-brand {

    color:
        #29241F;

    font-family:
        "Cormorant Garamond",
        Georgia,
        serif;

    font-size:
        31px;

    font-weight:
        600;

    letter-spacing:
        10px;

    margin-left:
        10px;
}


.footer-title {

    color:
        #786246;

    font-family:
        "Inter",
        Arial,
        sans-serif;

    font-size:
        9px;

    font-weight:
        600;

    letter-spacing:
        3px;

    margin-top:
        7px;
}


.footer-tagline {

    color:
        #A77E3F;

    font-family:
        "Cormorant Garamond",
        Georgia,
        serif;

    font-size:
        14px;

    font-style:
        italic;

    letter-spacing:
        2.5px;

    margin-top:
        12px;
}


.footer-bottom {

    color:
        #A39788;

    font-family:
        "Inter",
        Arial,
        sans-serif;

    font-size:
        8px;

    letter-spacing:
        1.3px;

    margin-top:
        18px;
}


/* ============================================================
   MOBILE
   ============================================================ */

@media (max-width: 700px) {

    .block-container {

        padding-left:
            1rem !important;

        padding-right:
            1rem !important;
    }


    .luxury-header {

        padding:
            30px 15px;

        margin-bottom:
            30px;
    }


    .bmw-brand {

        font-size:
            40px;

        letter-spacing:
            9px;

        margin-left:
            9px;
    }


    .bmw-subtitle {

        font-size:
            8px;

        letter-spacing:
            3px;
    }


    .bmw-tagline {

        font-size:
            14px;

        letter-spacing:
            1.5px;
    }


    .bmw-description {

        font-size:
            8px;

        letter-spacing:
            0.8px;
    }


    .section-title {

        font-size:
            14px;

        letter-spacing:
            2px;

        padding:
            13px 16px;
    }


    .stRadio > div {

        width:
            100% !important;

        gap:
            8px !important;
    }


    .stRadio > div > label {

        min-width:
            0 !important;

        flex:
            1 !important;

        padding:
            8px 10px !important;
    }


    .stRadio > div > label p {

        font-size:
            11px !important;
    }


    .footer-brand {

        font-size:
            24px;

        letter-spacing:
            7px;
    }
}

</style>
""",
    unsafe_allow_html=True
)


# ============================================================
# PDF EMPTY VALUE HANDLER
# ============================================================

def pdf_value(value):

    if value is None:
        return "-"

    try:

        if pd.isna(value):
            return "-"

    except Exception:
        pass

    value = str(value).strip()

    if value == "":
        return "-"

    return value


# ============================================================
# SAVE DATA TO EXCEL
# ============================================================

def save_to_excel(data):

    new_data = pd.DataFrame(
        [data]
    )

    if os.path.exists(EXCEL_FILE):

        try:

            old_data = pd.read_excel(
                EXCEL_FILE
            )

            final_data = pd.concat(
                [
                    old_data,
                    new_data
                ],
                ignore_index=True
            )

        except Exception:

            final_data = new_data

    else:

        final_data = new_data

    final_data.to_excel(
        EXCEL_FILE,
        index=False
    )


# ============================================================
# CREATE PDF
# ============================================================

def create_pdf(data):

    buffer = BytesIO()

    document = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=40,
        leftMargin=40,
        topMargin=40,
        bottomMargin=40
    )

    elements = []

    DARK = colors.HexColor(
        "#29241F"
    )

    GOLD = colors.HexColor(
        "#9C7A45"
    )

    BEIGE = colors.HexColor(
        "#E8DDCC"
    )

    LIGHT_BEIGE = colors.HexColor(
        "#FFF8EC"
    )

    title_style = ParagraphStyle(
        "BMWTitle",
        fontName=PDF_FONT,
        fontSize=30,
        leading=36,
        alignment=TA_CENTER,
        textColor=DARK,
        spaceAfter=4
    )

    subtitle_style = ParagraphStyle(
        "BMWSubtitle",
        fontName=PDF_FONT,
        fontSize=11,
        leading=16,
        alignment=TA_CENTER,
        textColor=GOLD,
        spaceAfter=12
    )

    tagline_style = ParagraphStyle(
        "BMWTagline",
        fontName=PDF_FONT,
        fontSize=13,
        leading=18,
        alignment=TA_CENTER,
        textColor=GOLD,
        spaceAfter=25
    )

    footer_style = ParagraphStyle(
        "Footer",
        fontName=PDF_FONT,
        fontSize=9,
        alignment=TA_CENTER,
        textColor=GOLD
    )

    header_text_style = ParagraphStyle(
        "TableHeader",
        fontName=PDF_FONT,
        fontSize=10,
        textColor=LIGHT_BEIGE,
        alignment=TA_CENTER
    )

    field_style = ParagraphStyle(
        "FieldStyle",
        fontName=PDF_FONT,
        fontSize=9,
        leading=13,
        textColor=DARK
    )

    value_style = ParagraphStyle(
        "ValueStyle",
        fontName=PDF_FONT,
        fontSize=9,
        leading=13,
        textColor=DARK
    )

    elements.append(
        Paragraph(
            "BMW",
            title_style
        )
    )

    elements.append(
        Paragraph(
            "VEHICLE QUOTATION REQUEST",
            subtitle_style
        )
    )

    header_line = Table(
        [[""]],
        colWidths=[180],
        rowHeights=[2]
    )

    header_line.setStyle(
        TableStyle(
            [
                (
                    "BACKGROUND",
                    (0, 0),
                    (-1, -1),
                    GOLD
                )
            ]
        )
    )

    elements.append(
        header_line
    )

    elements.append(
        Spacer(1, 12)
    )

    elements.append(
        Paragraph(
            "Driving Luxury. Delivering Excellence.",
            tagline_style
        )
    )

    table_data = [
        [
            Paragraph(
                "FIELD",
                header_text_style
            ),
            Paragraph(
                "DETAILS",
                header_text_style
            )
        ]
    ]

    for key, value in data.items():

        table_data.append(
            [
                Paragraph(
                    pdf_value(key),
                    field_style
                ),
                Paragraph(
                    pdf_value(value),
                    value_style
                )
            ]
        )

    table = Table(
        table_data,
        colWidths=[
            190,
            325
        ],
        repeatRows=1
    )

    table.setStyle(
        TableStyle(
            [
                (
                    "BACKGROUND",
                    (0, 0),
                    (-1, 0),
                    DARK
                ),
                (
                    "BACKGROUND",
                    (0, 1),
                    (0, -1),
                    BEIGE
                ),
                (
                    "BACKGROUND",
                    (1, 1),
                    (1, -1),
                    LIGHT_BEIGE
                ),
                (
                    "GRID",
                    (0, 0),
                    (-1, -1),
                    0.5,
                    GOLD
                ),
                (
                    "VALIGN",
                    (0, 0),
                    (-1, -1),
                    "MIDDLE"
                ),
                (
                    "LEFTPADDING",
                    (0, 0),
                    (-1, -1),
                    10
                ),
                (
                    "RIGHTPADDING",
                    (0, 0),
                    (-1, -1),
                    10
                ),
                (
                    "TOPPADDING",
                    (0, 0),
                    (-1, -1),
                    8
                ),
                (
                    "BOTTOMPADDING",
                    (0, 0),
                    (-1, -1),
                    8
                )
            ]
        )
    )

    elements.append(
        table
    )

    elements.append(
        Spacer(1, 25)
    )

    elements.append(
        Paragraph(
            "THE ULTIMATE DRIVING EXPERIENCE",
            footer_style
        )
    )

    elements.append(
        Spacer(1, 8)
    )

    elements.append(
        Paragraph(
            "Generated on: "
            +
            datetime.now().strftime(
                "%d-%m-%Y %I:%M %p"
            ),
            footer_style
        )
    )

    document.build(
        elements
    )

    buffer.seek(0)

    return buffer


# ============================================================
# LUXURY BMW PAGE HEADER
# ============================================================

st.markdown(
    """
<div class="luxury-header">

<div class="header-top-line"></div>

<div class="bmw-brand">
    BMW
</div>

<div class="bmw-subtitle">
    VEHICLE QUOTATION
</div>

<div class="header-divider">

<span></span>

<div class="divider-diamond">
◆
</div>

<span></span>

</div>

<div class="bmw-tagline">
THE ULTIMATE DRIVING EXPERIENCE
</div>

<div class="bmw-description">
Your journey to Sheer Driving Pleasure begins here.
</div>

</div>
""",
    unsafe_allow_html=True
)


# ============================================================
# FORM 1
# ============================================================

with st.form("bmw_main_form"):

    # ========================================================
    # CUSTOMER DETAILS
    # ========================================================

    st.markdown(
        '<div class="section-title">CUSTOMER DETAILS</div>',
        unsafe_allow_html=True
    )

    col1, col2 = st.columns(2)

    with col1:

        customer_name = st.text_input(
            "Customer Full Name *",
            placeholder="Enter customer full name"
        )

        mobile_number = st.text_input(
            "Mobile Number *",
            max_chars=10,
            placeholder="Enter 10 digit mobile number"
        )

        email = st.text_input(
            "Email Address",
            placeholder="Enter email address"
        )

    with col2:

        enquiry_date = st.date_input(
            "Enquiry Date",
            format="DD-MM-YYYY"
        )

        city = st.text_input(
            "City *",
            placeholder="Enter city"
        )

        customer_type = st.selectbox(
            "Customer Type *",
            [
                "Select customer type",
                "Individual",
                "Corporate",
                "Existing BMW Customer",
                "Other"
            ]
        )

    # ========================================================
    # VEHICLE REQUIREMENT
    # ========================================================

    st.markdown(
        '<div class="section-title">VEHICLE REQUIREMENT</div>',
        unsafe_allow_html=True
    )

    col1, col2 = st.columns(2)

    with col1:

        series = st.selectbox(
            "BMW Series / Category *",
            [
                "Select series / category",
                "2 Series",
                "3 Series",
                "4 Series",
                "5 Series",
                "7 Series",
                "8 Series",
                "X1",
                "X3",
                "X5",
                "X7",
                "XM",
                "iX1",
                "i4",
                "i5",
                "i7",
                "iX",
                "Other"
            ]
        )

        model = st.text_input(
            "Model Required *",
            placeholder="Enter preferred BMW model"
        )

        exterior_colour = st.text_input(
            "Preferred Exterior Colour",
            placeholder="e.g. Alpine White"
        )

    with col2:

        variant = st.text_input(
            "Preferred Variant",
            placeholder="Enter preferred variant"
        )

        interior_colour = st.text_input(
            "Preferred Interior / Upholstery",
            placeholder="Enter preferred interior"
        )

        quantity = st.number_input(
            "Number of Vehicles",
            min_value=1,
            value=1,
            step=1
        )

    # ========================================================
    # PURCHASE & FINANCE
    # ========================================================

    st.markdown(
        '<div class="section-title">PURCHASE & FINANCE DETAILS</div>',
        unsafe_allow_html=True
    )

    col1, col2 = st.columns(2)

    with col1:

        purchase_type = st.selectbox(
            "Purchase Type *",
            [
                "Select purchase type",
                "Cash Purchase",
                "Finance",
                "Lease",
                "Corporate Purchase"
            ]
        )

        finance_required = st.selectbox(
            "Finance Required?",
            [
                "Select option",
                "Yes",
                "No"
            ]
        )

    with col2:

        budget = st.text_input(
            "Approximate Budget",
            placeholder="e.g. ₹50,00,000"
        )

        expected_purchase = st.selectbox(
            "Expected Purchase Timeline",
            [
                "Select timeline",
                "Immediately",
                "Within 15 Days",
                "Within 1 Month",
                "Within 3 Months",
                "More Than 3 Months"
            ]
        )

    # ========================================================
    # CONTINUE BUTTON
    # ========================================================

    purchase_continue_clicked = st.form_submit_button(
        "CONTINUE",
        use_container_width=True
    )


# ============================================================
# CONTINUE PROCESSING
# ============================================================

if purchase_continue_clicked:

    if not customer_name.strip():

        st.warning(
            "Please enter the Customer Full Name before continuing."
        )

        st.session_state["purchase_continue"] = False

    elif not mobile_number.strip():

        st.warning(
            "Please enter the Mobile Number before continuing."
        )

        st.session_state["purchase_continue"] = False

    elif not mobile_number.isdigit():

        st.warning(
            "Mobile Number must contain numbers only."
        )

        st.session_state["purchase_continue"] = False

    elif len(mobile_number) != 10:

        st.warning(
            "Mobile Number must contain exactly 10 digits."
        )

        st.session_state["purchase_continue"] = False

    elif not city.strip():

        st.warning(
            "Please enter the City before continuing."
        )

        st.session_state["purchase_continue"] = False

    elif customer_type == "Select customer type":

        st.warning(
            "Please select Customer Type before continuing."
        )

        st.session_state["purchase_continue"] = False

    elif series == "Select series / category":

        st.warning(
            "Please select BMW Series / Category before continuing."
        )

        st.session_state["purchase_continue"] = False

    elif not model.strip():

        st.warning(
            "Please enter the required BMW Model before continuing."
        )

        st.session_state["purchase_continue"] = False

    elif purchase_type == "Select purchase type":

        st.warning(
            "Please select Purchase Type before continuing."
        )

        st.session_state["purchase_continue"] = False

    elif finance_required == "Select option":

        st.warning(
            "Please select whether Finance is required."
        )

        st.session_state["purchase_continue"] = False

    else:

        st.session_state["purchase_continue"] = True

        st.success(
            "Customer and vehicle details saved. You can continue below."
        )


# ============================================================
# CONTINUE STATUS
# ============================================================

if st.session_state["purchase_continue"]:

    st.markdown(
        """
        <div class="continue-status">
            ✓ DETAILS SAVED — PLEASE COMPLETE THE QUOTATION
        </div>
        """,
        unsafe_allow_html=True
    )


# ============================================================
# PHONE VALIDATION
# ============================================================

if mobile_number:

    if not mobile_number.isdigit():

        st.warning(
            "Mobile Number can contain numbers only."
        )


# ============================================================
# SCRATCH CARD SECTION
# ============================================================

st.markdown(
    '<div class="section-title">EXCLUSIVE BMW PRIVILEGE</div>',
    unsafe_allow_html=True
)

discount_amount = st.session_state["scratch_discount"]

discount_display = format_inr(
    discount_amount
)


# ============================================================
# SCRATCH CARD
# ============================================================

scratch_html = f"""

<!DOCTYPE html>

<html>

<head>

<meta name="viewport"
content="width=device-width, initial-scale=1.0">

<style>

* {{
    box-sizing:border-box;
}}

html,
body {{
    margin:0;
    padding:0;
    background:transparent;
    overflow:hidden;
}}

body {{
    font-family:
        Georgia,
        serif;
}}

.card {{

    position:relative;

    width:96%;

    max-width:700px;

    height:310px;

    margin:auto;

    border-radius:18px;

    overflow:hidden;

    background:
        radial-gradient(
            circle at 20% 20%,
            rgba(255,255,255,.12),
            transparent 20%
        ),

        radial-gradient(
            circle at 80% 70%,
            rgba(255,255,255,.08),
            transparent 25%
        ),

        linear-gradient(
            135deg,
            #211d19,
            #3b332a,
            #211d19
        );

    border:
        2px solid #C8A45D;

    box-shadow:
        0 15px 35px
        rgba(0,0,0,.30);
}}

.prize {{

    position:absolute;

    inset:0;

    display:flex;

    flex-direction:column;

    justify-content:center;

    align-items:center;

    text-align:center;

    padding:20px;
}}

.small {{

    color:#D4BB8A;

    font-size:13px;

    letter-spacing:5px;

    margin-bottom:12px;
}}

.title {{

    color:#F7EFE2;

    font-size:24px;

    font-weight:bold;

    letter-spacing:3px;

    margin-bottom:12px;
}}

.amount {{

    color:#D8B45C;

    font-size:48px;

    font-weight:bold;

    letter-spacing:2px;
}}

.note {{

    color:#D7CCBD;

    font-size:12px;

    margin-top:12px;

    letter-spacing:1px;

    line-height:1.5;
}}

canvas {{

    position:absolute;

    left:0;

    top:0;

    width:100%;

    height:100%;

    z-index:5;

    cursor:crosshair;

    touch-action:none;
}}

.confetti {{

    position:fixed;

    width:8px;

    height:14px;

    top:-20px;

    z-index:100;

    pointer-events:none;

    animation:
        fall 3s linear forwards;
}}

@keyframes fall {{

    0% {{

        transform:
            translateY(0)
            rotate(0deg);

        opacity:1;
    }}

    100% {{

        transform:
            translateY(100vh)
            rotate(720deg);

        opacity:0;
    }}

}}

</style>

</head>

<body>

<div class="card">

    <div class="prize">

        <div class="small">
            CONGRATULATIONS
        </div>

        <div class="title">
            YOUR EXCLUSIVE DISCOUNT
        </div>

        <div class="amount">
            {discount_display}
        </div>

        <div class="note">
            This exclusive BMW benefit will be
            included with your quotation request.
        </div>

    </div>

    <canvas id="canvas"></canvas>

</div>

<script>

const canvas =
    document.getElementById("canvas");

const ctx =
    canvas.getContext("2d");

const card =
    document.querySelector(".card");

let scratching = false;

let lastX = 0;

let lastY = 0;

let revealed = false;


function resize() {{

    const rect =
        card.getBoundingClientRect();

    const ratio =
        window.devicePixelRatio || 1;

    canvas.width =
        rect.width * ratio;

    canvas.height =
        rect.height * ratio;

    canvas.style.width =
        rect.width + "px";

    canvas.style.height =
        rect.height + "px";

    ctx.setTransform(
        ratio,
        0,
        0,
        ratio,
        0,
        0
    );

    drawSurface();
}}


function drawSurface() {{

    const w =
        card.clientWidth;

    const h =
        card.clientHeight;

    const gradient =
        ctx.createLinearGradient(
            0,
            0,
            w,
            h
        );

    gradient.addColorStop(
        0,
        "#999999"
    );

    gradient.addColorStop(
        0.25,
        "#E0E0E0"
    );

    gradient.addColorStop(
        0.5,
        "#AAAAAA"
    );

    gradient.addColorStop(
        0.75,
        "#D8D8D8"
    );

    gradient.addColorStop(
        1,
        "#888888"
    );

    ctx.fillStyle =
        gradient;

    ctx.fillRect(
        0,
        0,
        w,
        h
    );


    for (
        let i = 0;
        i < 120;
        i++
    ) {{

        const x =
            Math.random() * w;

        const y =
            Math.random() * h;

        ctx.fillStyle =
            "rgba(255,255,255,.12)";

        ctx.fillRect(
            x,
            y,
            Math.random() * 120,
            1
        );
    }}


    ctx.fillStyle =
        "#4B453D";

    ctx.textAlign =
        "center";

    ctx.font =
        "bold 20px Georgia";

    ctx.fillText(
        "SCRATCH TO REVEAL",
        w / 2,
        h / 2
    );

    ctx.font =
        "12px Georgia";

    ctx.fillText(
        "Reveal your exclusive BMW benefit",
        w / 2,
        h / 2 + 28
    );
}}


function position(event) {{

    const rect =
        canvas.getBoundingClientRect();

    let x;

    let y;

    if (
        event.touches &&
        event.touches.length
    ) {{

        x =
            event.touches[0].clientX;

        y =
            event.touches[0].clientY;

    }} else {{

        x =
            event.clientX;

        y =
            event.clientY;
    }}

    return {{

        x:
            x - rect.left,

        y:
            y - rect.top

    }};
}}


function start(event) {{

    if (revealed)
        return;

    scratching = true;

    const p =
        position(event);

    lastX = p.x;

    lastY = p.y;

    scratch(event);
}}


function scratch(event) {{

    if (
        !scratching ||
        revealed
    )
        return;

    event.preventDefault();

    const p =
        position(event);

    ctx.globalCompositeOperation =
        "destination-out";

    ctx.lineWidth =
        45;

    ctx.lineCap =
        "round";

    ctx.lineJoin =
        "round";

    ctx.beginPath();

    ctx.moveTo(
        lastX,
        lastY
    );

    ctx.lineTo(
        p.x,
        p.y
    );

    ctx.stroke();

    lastX =
        p.x;

    lastY =
        p.y;

    check();
}}


function stop() {{

    scratching = false;
}}


function check() {{

    const data =
        ctx.getImageData(
            0,
            0,
            canvas.width,
            canvas.height
        ).data;

    let clear = 0;

    for (
        let i = 3;
        i < data.length;
        i += 16
    ) {{

        if (data[i] === 0)
            clear++;
    }}

    const total =
        data.length / 16;

    const percentage =
        clear / total;

    if (percentage >= 0.50) {{

        reveal();
    }}
}}


function reveal() {{

    if (revealed)
        return;

    revealed = true;

    canvas.style.transition =
        "opacity .8s ease";

    canvas.style.opacity =
        "0";

    confetti();
}}


function confetti() {{

    const colors = [

        "#D4AF68",
        "#FFFFFF",
        "#C8A45D",
        "#9C7A45",
        "#E8DDCC"

    ];

    for (
        let i = 0;
        i < 100;
        i++
    ) {{

        const piece =
            document.createElement("div");

        piece.className =
            "confetti";

        piece.style.background =
            colors[
                Math.floor(
                    Math.random()
                    * colors.length
                )
            ];

        piece.style.left =
            Math.random()
            * 100 + "%";

        piece.style.animationDuration =
            2 +
            Math.random() * 2 +
            "s";

        piece.style.animationDelay =
            Math.random() * 0.7 +
            "s";

        document.body.appendChild(
            piece
        );

        setTimeout(
            () => piece.remove(),
            4500
        );
    }}
}}


canvas.addEventListener(
    "mousedown",
    start
);

canvas.addEventListener(
    "mousemove",
    scratch
);

canvas.addEventListener(
    "mouseup",
    stop
);

canvas.addEventListener(
    "mouseleave",
    stop
);

canvas.addEventListener(
    "touchstart",
    start,
    {{passive:false}}
);

canvas.addEventListener(
    "touchmove",
    scratch,
    {{passive:false}}
);

canvas.addEventListener(
    "touchend",
    stop
);

window.addEventListener(
    "resize",
    resize
);

setTimeout(
    resize,
    100
);

</script>

</body>

</html>

"""


components.html(
    scratch_html,
    height=330,
    scrolling=False
)


# ============================================================
# DISCOUNT INFORMATION
# ============================================================

st.markdown(
    f"""
    <div class="discount-info">
        EXCLUSIVE BMW PRIVILEGE &nbsp; ◆ &nbsp;
        DISCOUNT VALUE: <b>{discount_display}</b>
    </div>
    """,
    unsafe_allow_html=True
)


# ============================================================
# FORM 2
# ============================================================

with st.form("bmw_final_form"):

    # ========================================================
    # VEHICLE EXCHANGE
    # ========================================================

    st.markdown(
        '<div class="section-title">VEHICLE EXCHANGE</div>',
        unsafe_allow_html=True
    )

    exchange_required = st.radio(
        "Do you have a vehicle for exchange?",
        [
            "No",
            "Yes"
        ],
        horizontal=True
    )

    col1, col2 = st.columns(2)

    with col1:

        exchange_make = st.text_input(
            "Exchange Vehicle Make",
            placeholder="e.g. Mercedes-Benz"
        )

        exchange_model = st.text_input(
            "Exchange Vehicle Model",
            placeholder="Enter vehicle model"
        )

    with col2:

        exchange_year = st.text_input(
            "Year of Manufacture",
            placeholder="e.g. 2022"
        )

        exchange_registration = st.text_input(
            "Registration Number",
            placeholder="e.g. MH01AB1234"
        )

    # ========================================================
    # QUOTATION REQUIREMENTS
    # ========================================================

    st.markdown(
        '<div class="section-title">QUOTATION REQUIREMENTS</div>',
        unsafe_allow_html=True
    )

    quotation_required = st.multiselect(
        "Please include the following in the quotation",
        [
            "Ex-Showroom Price",
            "Registration Charges",
            "Insurance",
            "Accessories",
            "Extended Warranty",
            "BMW Service Package",
            "Finance / EMI Options",
            "Exchange Offer",
            "Corporate Offer",
            "Special Discount / Offer"
        ]
    )

    preferred_contact = st.selectbox(
        "Preferred Contact Method",
        [
            "Phone Call",
            "WhatsApp",
            "Email"
        ]
    )

    remarks = st.text_area(
        "Additional Requirements / Remarks",
        placeholder="Please enter any additional requirements..."
    )

    # ========================================================
    # DECLARATION
    # ========================================================

    st.markdown(
        '<div class="section-title">DECLARATION</div>',
        unsafe_allow_html=True
    )

    declaration = st.checkbox(
        "I confirm that the information provided above is correct."
    )

    submit = st.form_submit_button(
        "SUBMIT QUOTATION REQUEST",
        use_container_width=True
    )


# ============================================================
# PROCESS SUBMISSION
# ============================================================

if submit:

    if not st.session_state["purchase_continue"]:

        st.warning(
            "Please complete the first section and click CONTINUE before submitting."
        )

    elif not declaration:

        st.error(
            "Please confirm that the information provided is correct."
        )

    else:

        discount_value = format_inr(
            discount_amount
        )

        form_data = {

            "Submission Timestamp":
                datetime.now().strftime(
                    "%d-%m-%Y %I:%M:%S %p"
                ),

            "Customer Name":
                customer_name,

            "Mobile Number":
                mobile_number,

            "Email":
                email,

            "Enquiry Date":
                enquiry_date.strftime(
                    "%d-%m-%Y"
                ),

            "City":
                city,

            "Customer Type":
                customer_type,

            "BMW Series":
                series,

            "Model":
                model,

            "Variant":
                variant,

            "Exterior Colour":
                exterior_colour,

            "Interior Colour":
                interior_colour,

            "Quantity":
                quantity,

            "Purchase Type":
                purchase_type,

            "Finance Required":
                finance_required,

            "Approximate Budget":
                budget,

            "Expected Purchase Timeline":
                expected_purchase,

            "Exclusive Scratch Card Discount":
                discount_value,

            "Exchange Required":
                exchange_required,

            "Exchange Make":
                exchange_make,

            "Exchange Model":
                exchange_model,

            "Exchange Year":
                exchange_year,

            "Exchange Registration":
                exchange_registration,

            "Quotation Requirements":
                ", ".join(
                    quotation_required
                ),

            "Preferred Contact Method":
                preferred_contact,

            "Additional Remarks":
                remarks
        }

        try:

            # =================================================
            # SAVE TO EXCEL
            # =================================================

            save_to_excel(
                form_data
            )

            # =================================================
            # CREATE PDF
            # =================================================

            pdf_file = create_pdf(
                form_data
            )

            st.session_state[
                "pdf_file"
            ] = pdf_file.getvalue()

            # =================================================
            # SAFE CUSTOMER NAME
            # =================================================

            safe_customer_name = re.sub(
                r"[^A-Za-z0-9_-]",
                "_",
                customer_name
            )

            st.session_state[
                "pdf_name"
            ] = (
                "BMW_Quotation_"
                +
                safe_customer_name
                +
                ".pdf"
            )

            st.success(
                "Your BMW quotation request has been submitted successfully."
            )

        except Exception as e:

            st.error(
                f"An error occurred: {e}"
            )


# ============================================================
# DOWNLOAD PDF
# ============================================================

if st.session_state["pdf_file"] is not None:

    st.markdown(
        "### Your Quotation Request"
    )

    st.download_button(
        label="DOWNLOAD AS PDF",
        data=st.session_state["pdf_file"],
        file_name=st.session_state["pdf_name"],
        mime="application/pdf",
        use_container_width=True
    )


# ============================================================
# LUXURY BMW FOOTER
# ============================================================

st.markdown(
    """
<div class="luxury-footer">

<div class="footer-divider">

<span></span>

<div class="footer-diamond">
    ◆
</div>

<span></span>

</div>

<div class="footer-brand">
    BMW
</div>

<div class="footer-title">
    VEHICLE QUOTATION REQUEST
</div>

<div class="footer-tagline">
    SHEER DRIVING PLEASURE
</div>

<div class="footer-bottom">
    © BMW &nbsp; | &nbsp; CUSTOMER QUOTATION SERVICES
</div>

</div>
""",
    unsafe_allow_html=True
)
