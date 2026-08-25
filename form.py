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

EXCEL_FILE = r"C:\Users\mitanshu.patil\OneDrive - INFINITY CARS PRIVATE LIMITED\Desktop\Dashboard\quotation_responses.xlsx"

st.set_page_config(
    page_title="BMW Vehicle Quotation Request",
    page_icon="🚘",
    layout="centered"
)


# ============================================================
# SESSION STATE
# ============================================================

# Random BMW discount
#
# Discount will be generated in ₹5,000 increments
# between ₹25,000 and ₹1,00,000.
#
# Examples:
# ₹25,000
# ₹30,000
# ₹35,000
# ...
# ₹1,00,000

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


# FIXED CONTINUE STATE

if "purchase_continue" not in st.session_state:

    st.session_state["purchase_continue"] = False


# ============================================================
# INDIAN RUPEE FORMATTER
# ============================================================

def format_inr(amount):
    """
    Format a number using the Indian numbering system.

    Examples:

    25000     -> ₹25,000
    50000     -> ₹50,000
    100000    -> ₹1,00,000
    1250000   -> ₹12,50,000
    10000000  -> ₹1,00,00,000
    """

    try:

        amount = int(float(amount))

    except Exception:

        return "₹0"


    amount_str = str(abs(amount))


    # Numbers below 1,000

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
# REPORTLAB UNICODE FONT
# ============================================================

def register_pdf_font():

    """
    Register a Unicode font that supports the Indian Rupee
    symbol (₹).

    DejaVu Sans is commonly available on Streamlit Cloud,
    Linux and many Windows installations.
    """

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


    # Try system font locations dynamically

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


    # Final fallback

    return "Times-Roman"


# ============================================================
# REGISTER PDF FONT
# ============================================================

PDF_FONT = register_pdf_font()


# ============================================================
# LUXURY DESIGN
# ============================================================

st.markdown(
    """
    <style>

    /* ========================================================
       MAIN PAGE
       ======================================================== */

    .stApp {

        background-color: #E8DDCC;

        color: #29241F;
    }


    .block-container {

        max-width: 1000px;

        padding-top: 1.5rem;

        padding-bottom: 3rem;
    }


    #MainMenu {

        visibility: hidden;
    }


    footer {

        visibility: hidden !important;
    }


    /* ========================================================
       GENERAL FONT
       ======================================================== */

    html,
    body,
    [class*="css"] {

        font-family:
            Georgia,
            "Times New Roman",
            serif;
    }


    /* ========================================================
       LUXURY HEADER
       ======================================================== */

    .luxury-header {

        position: relative;

        text-align: center;

        margin-top: 5px;

        margin-bottom: 38px;

        padding: 34px 25px 32px 25px;

        background:
            linear-gradient(
                145deg,
                #211D19 0%,
                #302A23 45%,
                #1D1916 100%
            );

        border-radius: 7px;

        border: 1px solid #6E5A3C;

        box-shadow:
            0 14px 35px
            rgba(0,0,0,0.22);

        overflow: hidden;
    }


    .luxury-header::before {

        content: "";

        position: absolute;

        top: 7px;

        left: 7px;

        right: 7px;

        bottom: 7px;

        border: 1px solid
        rgba(200,164,93,0.25);

        border-radius: 4px;

        pointer-events: none;
    }


    .header-top-line {

        position: absolute;

        top: 0;

        left: 0;

        right: 0;

        height: 3px;

        background:
            linear-gradient(
                90deg,
                transparent,
                #C8A45D,
                #F0D28B,
                #C8A45D,
                transparent
            );
    }


    /* Premium BMW wordmark */

    .bmw-brand {

        color: #F7F0E5;

        font-family:
            "Bodoni 72",
            Didot,
            "Bodoni MT",
            "Times New Roman",
            serif;

        font-size: 47px;

        font-weight: 600;

        letter-spacing: 12px;

        line-height: 1;

        margin-left: 12px;

        text-shadow:
            0 2px 8px
            rgba(0,0,0,0.45);
    }


    /* Header subtitle */

    .bmw-subtitle {

        color: #C8A45D;

        font-family:
            "Bodoni 72",
            Didot,
            "Bodoni MT",
            "Times New Roman",
            serif;

        font-size: 12px;

        font-weight: 600;

        letter-spacing: 5px;

        margin-top: 14px;
    }


    /* Header divider */

    .header-divider {

        display: flex;

        align-items: center;

        justify-content: center;

        gap: 13px;

        max-width: 310px;

        margin:
            19px auto
            16px auto;
    }


    .header-divider span {

        height: 1px;

        flex: 1;

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

        color: #D2B46A;

        font-size: 8px;
    }


    .bmw-tagline {

        color: #EFE5D7;

        font-family:
            "Baskerville",
            "Bodoni 72",
            Didot,
            Georgia,
            serif;

        font-size: 15px;

        font-style: italic;

        letter-spacing: 3px;
    }


    .bmw-description {

        color: #BBAE9B;

        font-family:
            Georgia,
            serif;

        font-size: 10px;

        letter-spacing: 1.4px;

        margin-top: 10px;
    }


    /* ========================================================
       LABELS
       ======================================================== */

    .stTextInput label,
    .stSelectbox label,
    .stTextArea label,
    .stDateInput label,
    .stNumberInput label,
    .stRadio label,
    .stMultiSelect label,
    .stCheckbox label {

        font-family:
            Georgia,
            "Times New Roman",
            serif !important;

        font-weight: bold !important;

        color: #3A3128 !important;
    }


    /* ========================================================
       INPUTS
       ======================================================== */

    .stTextInput > div > div,
    .stNumberInput > div > div,
    .stDateInput > div > div,
    .stTextArea > div > div {

        background-color: #FFFDF8 !important;

        border: none !important;

        outline: none !important;

        box-shadow: none !important;

        border-radius: 5px !important;
    }


    .stTextInput input,
    .stNumberInput input,
    .stDateInput input,
    .stTextArea textarea {

        background-color: #FFFDF8 !important;

        border: none !important;

        outline: none !important;

        box-shadow: none !important;

        color: #29241F !important;

        font-family:
            Georgia,
            "Times New Roman",
            serif !important;
    }


    .stTextInput > div > div:focus-within,
    .stNumberInput > div > div:focus-within,
    .stDateInput > div > div:focus-within,
    .stTextArea > div > div:focus-within {

        border: none !important;

        outline: none !important;

        box-shadow: none !important;
    }


    /* ========================================================
       SELECT BOX
       ======================================================== */

    .stSelectbox div[data-baseweb="select"],
    .stMultiSelect div[data-baseweb="select"] {

        background-color: #FFFDF8 !important;

        border: none !important;

        outline: none !important;

        box-shadow: none !important;

        border-radius: 5px !important;

        font-family:
            Georgia,
            "Times New Roman",
            serif !important;
    }


    /* ========================================================
       SECTION TITLES
       ======================================================== */

    .section-title {

        background:
            linear-gradient(
                90deg,
                #29241F,
                #383027,
                #29241F
            );

        color: #F5EBDD;

        padding: 13px 20px;

        font-family:
            "Bodoni 72",
            Didot,
            "Bodoni MT",
            Georgia,
            serif;

        font-size: 16px;

        font-weight: 600;

        letter-spacing: 2.5px;

        border-left: 4px solid #C8A45D;

        margin-top: 30px;

        margin-bottom: 18px;

        border-radius: 4px;

        box-shadow:
            0 4px 12px
            rgba(0,0,0,0.10);
    }


    /* ========================================================
       BUTTONS
       ======================================================== */

    .stButton > button {

        background-color: #9C7A45 !important;

        color: white !important;

        border: none !important;

        border-radius: 5px !important;

        height: 50px !important;

        font-family:
            "Bodoni 72",
            Didot,
            Georgia,
            serif !important;

        font-size: 13px !important;

        font-weight: bold !important;

        letter-spacing: 2px !important;

        transition:
            all 0.25s ease !important;
    }


    .stButton > button:hover {

        background-color: #29241F !important;

        color: white !important;

        transform:
            translateY(-1px);
    }


    /* ========================================================
       FORM SUBMIT BUTTON
       ======================================================== */

    .stFormSubmitButton button {

        background:
            linear-gradient(
                135deg,
                #29241F,
                #3A3128
            ) !important;

        color: #F5EBDD !important;

        border: 1px solid #9C7A45 !important;

        border-radius: 5px !important;

        height: 55px !important;

        font-family:
            "Bodoni 72",
            Didot,
            Georgia,
            serif !important;

        font-size: 14px !important;

        font-weight: bold !important;

        letter-spacing: 2.5px !important;

        transition:
            all 0.25s ease !important;
    }


    .stFormSubmitButton button:hover {

        background:
            linear-gradient(
                135deg,
                #9C7A45,
                #B89350
            ) !important;

        color: white !important;
    }


    /* ========================================================
       DOWNLOAD BUTTON
       ======================================================== */

    .stDownloadButton button {

        background:
            linear-gradient(
                135deg,
                #9C7A45,
                #B18C50
            ) !important;

        color: white !important;

        border: none !important;

        border-radius: 5px !important;

        height: 55px !important;

        font-family:
            "Bodoni 72",
            Didot,
            Georgia,
            serif !important;

        font-size: 14px !important;

        font-weight: bold !important;

        letter-spacing: 2px !important;
    }


    /* ========================================================
       CONTINUE SUCCESS INDICATOR
       ======================================================== */

    .continue-status {

        text-align: center;

        color: #79613D;

        font-family:
            Georgia,
            serif;

        font-size: 11px;

        letter-spacing: 2px;

        margin-top: 8px;

        margin-bottom: 4px;
    }


    /* ========================================================
       DISCOUNT INFORMATION
       ======================================================== */

    .discount-info {

        text-align: center;

        color: #79613D;

        font-family:
            Georgia,
            serif;

        font-size: 10px;

        letter-spacing: 1.5px;

        margin-top: 10px;
    }


    /* ========================================================
       LUXURY FOOTER
       ======================================================== */

    .luxury-footer {

        text-align: center;

        margin-top: 60px;

        margin-bottom: 15px;

        padding:
            28px 20px
            20px 20px;

        font-family:
            Georgia,
            serif;
    }


    .footer-divider {

        display: flex;

        align-items: center;

        justify-content: center;

        gap: 13px;

        max-width: 280px;

        margin:
            0 auto
            22px auto;
    }


    .footer-divider span {

        height: 1px;

        flex: 1;

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

        color: #9C7A45;

        font-size: 8px;
    }


    .footer-brand {

        color: #29241F;

        font-family:
            "Bodoni 72",
            Didot,
            "Bodoni MT",
            "Times New Roman",
            serif;

        font-size: 26px;

        font-weight: 600;

        letter-spacing: 8px;

        margin-left: 8px;
    }


    .footer-title {

        color: #786246;

        font-family:
            "Bodoni 72",
            Didot,
            "Bodoni MT",
            Georgia,
            serif;

        font-size: 10px;

        font-weight: 600;

        letter-spacing: 3.5px;

        margin-top: 8px;
    }


    .footer-tagline {

        color: #9C7A45;

        font-family:
            Baskerville,
            Georgia,
            serif;

        font-size: 11px;

        font-style: italic;

        letter-spacing: 2.5px;

        margin-top: 13px;
    }


    .footer-bottom {

        color: #A39788;

        font-family:
            Georgia,
            serif;

        font-size: 8px;

        letter-spacing: 1.5px;

        margin-top: 18px;
    }


    /* ========================================================
       MOBILE
       ======================================================== */

    @media (max-width: 600px) {

        .luxury-header {

            padding:
                28px 12px
                28px 12px;

            margin-bottom: 30px;
        }


        .bmw-brand {

            font-size: 38px;

            letter-spacing: 8px;

            margin-left: 8px;
        }


        .bmw-subtitle {

            font-size: 9px;

            letter-spacing: 3px;
        }


        .bmw-tagline {

            font-size: 12px;

            letter-spacing: 2px;
        }


        .bmw-description {

            font-size: 8px;

            letter-spacing: 1px;
        }


        .section-title {

            font-size: 14px;

            letter-spacing: 2px;
        }


        .footer-brand {

            font-size: 22px;

            letter-spacing: 6px;
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


    # ========================================================
    # PDF STYLES
    # ========================================================

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


    # ========================================================
    # PDF HEADER
    # ========================================================

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


    # ========================================================
    # PDF TABLE
    # ========================================================

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
            "Customer Full Name *"
        )


        mobile_number = st.text_input(
            "Mobile Number *",
            max_chars=10,
            placeholder="Enter 10 digit mobile number"
        )


        email = st.text_input(
            "Email Address"
        )


    with col2:

        enquiry_date = st.date_input(
            "Enquiry Date"
        )


        city = st.text_input(
            "City *"
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
            "Model Required *"
        )


        exterior_colour = st.text_input(
            "Preferred Exterior Colour"
        )


    with col2:

        variant = st.text_input(
            "Preferred Variant"
        )


        interior_colour = st.text_input(
            "Preferred Interior / Upholstery"
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
            "Approximate Budget"
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
# CONTINUE BUTTON PROCESSING
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
# SCRATCH CARD
# ============================================================

st.markdown(
    '<div class="section-title">EXCLUSIVE BMW PRIVILEGE</div>',
    unsafe_allow_html=True
)


discount_amount = st.session_state["scratch_discount"]


# ============================================================
# FORMATTED INR DISCOUNT
# ============================================================

discount_display = format_inr(
    discount_amount
)


# ============================================================
# SCRATCH CARD HTML
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

html, body {{
    margin:0;
    padding:0;
    background:transparent;
    overflow:hidden;
}}

body {{
    font-family:Georgia,serif;
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

    border:2px solid #C8A45D;

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

    animation:fall 3s linear forwards;
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
            "Exchange Vehicle Make"
        )


        exchange_model = st.text_input(
            "Exchange Vehicle Model"
        )


    with col2:

        exchange_year = st.text_input(
            "Year of Manufacture"
        )


        exchange_registration = st.text_input(
            "Registration Number"
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
        "Additional Requirements / Remarks"
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

        # ====================================================
        # FINAL DISCOUNT FORMAT
        # ====================================================

        discount_value = format_inr(
            discount_amount
        )


        # ====================================================
        # FORM DATA
        # ====================================================

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


            # =================================================
            # FORMATTED INDIAN RUPEE VALUE
            # =================================================

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
            # SAVE EXCEL
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
