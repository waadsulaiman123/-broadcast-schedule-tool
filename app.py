"""
app.py — نظام إدارة جداول البث (النسخة الأولى الحقيقية)
==========================================================
هذا نفس منطق step2 (محرك الترحيل) وstep3 (توليد الجدول) اللي بنيناهم
قبل، لكن هذي المرة ملفوفين بواجهة ويب حقيقية بواسطة Streamlit.

تشغيل هذا الملف:
    streamlit run app.py
    (أو لو ما اشتغل: python -m streamlit run app.py)

بعدها بيفتح لك تلقائيًا تبويب جديد في المتصفح فيه التطبيق شغّال.
"""

import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
import io
import re
import base64
from pathlib import Path

# ============================================================
# إعدادات الصفحة + هوية التعليمية (Brand Guidelines 2026)
# ============================================================
st.set_page_config(page_title="نظام إدارة جداول البث", page_icon="📅", layout="wide")

# نحوّل الشعار لصيغة base64 عشان نقدر نضمّه داخل HTML مباشرة
_logo_path = Path(__file__).parent / "assets" / "header_logo_crop.png"
_logo_b64 = base64.b64encode(_logo_path.read_bytes()).decode() if _logo_path.exists() else ""

st.markdown("""
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Cairo:wght@700;800&family=Tajawal:wght@400;500;700&display=swap" rel="stylesheet">
<style>
    :root{
        --navy:#22577B; --teal-dark:#2EA1A1; --teal-light:#C1DFC4;
        --green:#90C685; --aqua:#6CBE99; --gray:#575756;
        --ok:#2E8B67; --ok-bg:#E7F5EE; --warn:#B8860B; --warn-bg:#FBF3DF;
        --crit:#C0392B; --crit-bg:#FBEAE8;
    }
    html, body, .stApp, [data-testid="stAppViewContainer"],
    [data-testid="stVerticalBlock"], [data-testid="stHorizontalBlock"],
    [data-testid="stMarkdownContainer"], [data-testid="stMetric"],
    [data-testid="stExpander"], [data-testid="stAlert"],
    .stMarkdown, .stText, .stAlert, .stCheckbox, .stSelectbox, .stButton, .stRadio {
        direction: rtl !important;
        text-align: right !important;
        font-family: 'Tajawal', sans-serif !important;
    }
    [data-testid="stMetricLabel"], [data-testid="stMetricValue"] {
        direction: rtl !important; text-align: right !important; width: 100% !important;
    }
    h1, h2, h3 { font-family: 'Cairo', sans-serif !important; color: var(--navy) !important; }
    table { direction: rtl !important; text-align: right !important; }
    th, td { text-align: right !important; }
    thead tr th { background: var(--teal-light) !important; color: var(--navy) !important; font-weight: 700 !important; }

    /* شريط العنوان بهوية التعليمية */
    .talemia-topbar{
        background: var(--navy); border-radius: 16px; padding: 16px 22px; margin-bottom: 22px;
        display:flex; align-items:center; justify-content:space-between; position:relative; overflow:hidden;
    }
    .talemia-topbar::after{
        content:""; position:absolute; inset:0; opacity:.5; pointer-events:none;
        background: radial-gradient(circle at 100% 0%, transparent 55%, rgba(108,190,153,.18) 56%, transparent 62%);
    }
    .talemia-logo-pill{ background:#fff; border-radius:10px; padding:5px 10px; display:flex; align-items:center; z-index:2; }
    .talemia-logo-pill img{ height:30px; display:block; }
    .talemia-title-block{ display:flex; align-items:center; gap:14px; z-index:2; }
    .talemia-title-block .t1{ font-family:'Cairo',sans-serif; font-weight:800; font-size:19px; color:#fff; margin:0; }
    .talemia-title-block .t2{ font-size:12.5px; color:var(--teal-light); margin:0; }

    /* بطاقات إحصائية بهوية التعليمية */
    .stat-row{ display:flex; gap:12px; margin-bottom:8px; flex-wrap:wrap; }
    .stat-card{ flex:1; min-width:150px; background:#fff; border:1px solid #E2ECE8; border-radius:12px; padding:14px 16px; border-top:3px solid var(--navy); }
    .stat-card.ok{ border-top-color: var(--aqua); }
    .stat-card.warn{ border-top-color: var(--warn); }
    .stat-card.crit{ border-top-color: var(--crit); }
    .stat-card .num{ font-family:'Cairo',sans-serif; font-weight:800; font-size:26px; color:var(--navy); }
    .stat-card .lbl{ font-size:12px; color:var(--gray); margin-top:4px; }

    /* جديد: بطاقات خطوات مرقّمة بهوية التعليمية */
    .step-card{
        background:#fff; border:1px solid #E2ECE8; border-radius:16px;
        padding:20px 24px; margin-bottom:20px; box-shadow:0 2px 10px rgba(34,87,123,.05);
    }
    .step-card .step-head{ display:flex; align-items:center; gap:12px; margin-bottom:4px; }
    .step-card .step-num{
        background:var(--navy); color:#fff; width:34px; height:34px; border-radius:10px;
        display:flex; align-items:center; justify-content:center; font-family:'Cairo',sans-serif;
        font-weight:800; font-size:16px; flex-shrink:0;
    }
    .step-card .step-title{ font-family:'Cairo',sans-serif; font-weight:800; font-size:18px; color:var(--navy); margin:0; }
    .step-card .step-sub{ font-size:13px; color:var(--gray); margin:6px 0 16px 46px; line-height:1.7; }

    /* جديد: تنبيهات مخصّصة بهوية متسقة بدل صناديق ستريم لِت الافتراضية */
    .result-banner{ border-radius:12px; padding:14px 18px; margin:10px 0; font-size:14.5px; font-weight:600; display:flex; align-items:center; gap:10px; }
    .result-banner.ok{ background:var(--ok-bg); color:var(--ok); border:1px solid #BFE3D2; }
    .result-banner.warn{ background:var(--warn-bg); color:#8a6a13; border:1px solid #F0DFA8; }
    .result-banner.info{ background:#EAF2F8; color:var(--navy); border:1px solid #CBDDEA; }

    /* جديد: تحسين شكل أزرار الرفع الأساسية */
    .stButton > button{
        border-radius:12px !important; font-family:'Cairo',sans-serif !important; font-weight:800 !important;
        padding:10px 22px !important; box-shadow:0 3px 10px rgba(34,87,123,.12) !important;
        transition: transform .12s ease !important;
    }
    .stButton > button:hover{ transform: translateY(-1px); }
    .stButton > button[kind="primary"]{ background: var(--navy) !important; border-color: var(--navy) !important; }

    /* جديد: تمييز الفصل الدراسي المختار بشكل أوضح */
    div[role="radiogroup"] label{
        border:1.5px solid #E2ECE8 !important; border-radius:10px !important; padding:6px 18px !important;
        margin-inline-start:6px !important; transition: all .15s ease !important;
    }

    /* جديد: خلفية عامة هادئة وفواصل أنظف */
    [data-testid="stAppViewContainer"] > .main{ background: #F7FAF9; }
    hr{ margin:28px 0 !important; border-color:#E2ECE8 !important; }
</style>
""", unsafe_allow_html=True)

if _logo_b64:
    st.markdown(f"""
    <div class="talemia-topbar">
        <div class="talemia-title-block">
            <div>
                <p class="t1">📅 نظام إدارة جداول البث</p>
            </div>
        </div>
        <div class="talemia-logo-pill"><img src="data:image/png;base64,{_logo_b64}"></div>
    </div>
    """, unsafe_allow_html=True)

DAYS = ["الأحد", "الاثنين", "الثلاثاء", "الأربعاء", "الخميس"]


# ============================================================
# جديد: توحيد أسماء الأعمدة — يخلي البرنامج يقبل ملفك الحقيقي
# بأسمائه الأصلية (زي "المقرر/المادة" و"ترتيب بث الدروس")
# بدون ما تغيّرين شي فيه يدويًا
#
# جديد: بدل قائمة أسماء مطابقة حرفيًا (كانت تنكسر لو تغيّرت صياغة العمود
# شوي)، نستخدم الآن مطابقة بالكلمات المفتاحية — أي عمود يحتوي هالكلمات
# (بأي ترتيب أو صياغة) يتحوّل تلقائيًا للاسم الداخلي المطلوب.
# ============================================================
COLUMN_KEYWORD_RULES = [
    (["مقرر"], "المادة"),
    (["ماده"], "المادة"),  # يغطي أيضًا "مادة" لأن الألف بعد الميم موجودة أصلًا في الكلمة
    (["ترتيب", "بث"], "ترتيب_البث"),
    (["عنوان", "درس"], "عنوان_الدرس"),
    (["يوتيوب"], "رابط_يوتيوب"),
    (["فصل", "دراسي"], "الفصل_الدراسي"),
    (["اسبوع"], "الأسبوع"),  # يمسك "الاسبوع" بدون همزة و"الأسبوع" بالهمزة
]

# جديد: بعض الملفات تكتب رقم الأسبوع/الحصة بالكلمات العربية بدل الأرقام
# (مثلًا "الأول" بدل 1، "الخامسة" بدل 5) — هذا القاموسان يحوّلونها لأرقام
WEEK_WORD_TO_NUM = {
    "الاول": 1, "الثاني": 2, "الثالث": 3, "الرابع": 4, "الخامس": 5,
    "السادس": 6, "السابع": 7, "الثامن": 8, "التاسع": 9, "العاشر": 10,
    "الحادي عشر": 11, "الثاني عشر": 12, "الثالث عشر": 13, "الرابع عشر": 14,
    "الخامس عشر": 15, "السادس عشر": 16, "السابع عشر": 17, "الثامن عشر": 18,
    "التاسع عشر": 19, "العشرون": 20,
}
PERIOD_WORD_TO_NUM = {
    "الاولي": 1, "الثانيه": 2, "الثالثه": 3, "الرابعه": 4,
    "الخامسه": 5, "السادسه": 6, "السابعه": 7,
}
# جديد: توحيد أسماء الأيام (بعض الملفات تكتب "الأثنين" بالهمزة، والكود
# الداخلي يتوقع "الاثنين" بدون همزة بالضبط)
DAY_NAME_CANON = {
    "الاحد": "الأحد",
    "الاثنين": "الاثنين",
    "الثلاثاء": "الثلاثاء",
    "الاربعاء": "الأربعاء",
    "الخميس": "الخميس",
}


def _normalize_arabic(text: str) -> str:
    """توحيد بسيط للنص العربي عشان المطابقة ما تفشل بسبب اختلاف الألف/التاء المربوطة."""
    text = str(text).strip()
    for a, b in [("أ", "ا"), ("إ", "ا"), ("آ", "ا"), ("ة", "ه"), ("ى", "ي")]:
        text = text.replace(a, b)
    return text


def _convert_word_number(value, word_map: dict):
    """يحوّل قيمة (رقم أو كلمة عربية) لرقم صحيح. يرجّع القيمة الأصلية لو ما قدر يحوّلها."""
    if pd.isna(value):
        return value
    # لو كانت أصلًا رقمًا (أو نص رقمي)، رجّعيها كرقم صحيح مباشرة
    try:
        return int(float(value))
    except (ValueError, TypeError):
        pass
    # وإلا جرّبي تطابقها ككلمة عربية بعد التطبيع
    norm = _normalize_arabic(value)
    return word_map.get(norm, value)  # لو ما لقيناها، نرجّع القيمة الأصلية كما هي


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]  # إزالة أي مسافات زائدة بالاسم

    # جديد: نفحص كل قاعدة على حدة (مو كل عمود على حدة)، ولو فيه أكثر من عمود
    # يطابق نفس القاعدة، نختار الأقصر اسمًا دائمًا — هذا يمنع عمود شرح طويل
    # (فيه كل الكلمات المفتاحية بالصدفة داخل جملة توضيحية) من "خطف" مكان
    # عمود قصير حقيقي يطابق نفس الكلمات بشكل مباشر وأوضح
    #
    # استثناء: عمود "الفصل الدراسي" — ملفاتك الحقيقية فيها غالبًا عمودين
    # يطابقون هالكلمات: العمود العام (نص كامل "الفصل الدراسي الأول"، وأحيانًا
    # ناقص/غلط بالعنوان) والعمود التفصيلي لكل درس (فيه "-1أو 2-" برأسه ورقم
    # نظيف بالبيانات). نفضّل الثاني دائمًا لأنه أوثق وأنظف.
    rename_map = {}
    claimed_cols = set()
    for keywords, target in COLUMN_KEYWORD_RULES:
        candidates = [
            c for c in df.columns
            if c not in claimed_cols
            and all(_normalize_arabic(k) in _normalize_arabic(c) for k in keywords)
        ]
        if not candidates:
            continue

        if target == "الفصل_الدراسي":
            # نفضّل العمود اللي برأسه رقم (زي "-1أو 2-") — هذا هو عمود
            # الفصل الدراسي "لكل درس"، أدق وأثبت من العمود العام
            marked = [c for c in candidates if any(ch.isdigit() for ch in c)]
            candidates = marked if marked else candidates

        candidates.sort(key=len)  # الأقصر أولًا = الأقرب لتطابق حرفي مباشر
        chosen = candidates[0]
        rename_map[chosen] = target
        claimed_cols.add(chosen)

    df = df.rename(columns=rename_map)

    # جديد: نتجاهل الصفوف الفاضية تمامًا (كل الأعمدة الأساسية فاضية) —
    # هذي عادة أسطر فاضية بالصدفة داخل الإكسل (مسافة بصرية بين أقسام مثلًا)
    # مو أخطاء بيانات حقيقية، فما نزعج المستخدمة بتقرير عنها.
    _core_cols = [c for c in ["الأسبوع", "الحصة", "اليوم", "عنوان_الدرس"] if c in df.columns]
    if _core_cols:
        _all_blank = df[_core_cols].isna().all(axis=1)
        n_blank = _all_blank.sum()
        if n_blank > 0:
            df = df[~_all_blank].reset_index(drop=True)

    # توحيد قيم الفصل الدراسي — تدعم الحالتين:
    # (أ) نص كامل "الفصل الدراسي الأول" → نشيل الجزء الزايد
    # (ب) رقم نظيف 1 أو 2 (من عمود "-1أو 2-") → نحوّله لكلمة "الأول"/"الثاني"
    if "الفصل_الدراسي" in df.columns:
        def _clean_semester(v):
            if pd.isna(v):
                return v
            try:
                n = int(float(v))
                return {1: "الأول", 2: "الثاني"}.get(n, str(v))
            except (ValueError, TypeError):
                pass
            v = str(v).strip()
            v = v.replace("الفصل الدراسي", "").strip()
            return v

        df["الفصل_الدراسي"] = df["الفصل_الدراسي"].apply(_clean_semester)

    # جديد: تحويل الأسبوع والحصة من كلمات عربية لأرقام صحيحة (لو لزم الأمر)
    if "الأسبوع" in df.columns:
        df["الأسبوع"] = df["الأسبوع"].apply(lambda v: _convert_word_number(v, WEEK_WORD_TO_NUM))
    if "الحصة" in df.columns:
        df["الحصة"] = df["الحصة"].apply(lambda v: _convert_word_number(v, PERIOD_WORD_TO_NUM))

    # جديد: توحيد أسماء الأيام (همزة/بدون همزة)
    if "اليوم" in df.columns:
        df["اليوم"] = df["اليوم"].apply(
            lambda v: DAY_NAME_CANON.get(_normalize_arabic(v), v) if pd.notna(v) else v
        )

    # ملاحظة: كنا نستبعد دروس بدون رقم ترتيب حقيقي (زي "_" بدل رقم) — بناءً
    # على طلب المستخدمة، صار الآن ما نستبعدهم، يدخلون جدول البث عادي.
    # (الدالة _drop_unordered_rows محفوظة تحت لو احتجناها ثانية.)

    return df


def _drop_unordered_rows(df: pd.DataFrame) -> pd.DataFrame:
    def _is_numeric(v):
        try:
            float(v)
            return True
        except (ValueError, TypeError):
            return False

    # مهم: .astype(bool) صراحة — لو df فاضي (٠ صف)، .apply() على عمود فاضي
    # يرجّع نوع بيانات object مو bool، وهذا يخلي df[is_num] يمسح الأعمدة
    # كلها بالغلط بدل ما يفلتر الصفوف بس. هذا إصلاح حرج لأي ملف/قالب فاضي.
    is_num = df["ترتيب_البث"].apply(_is_numeric).astype(bool)
    n_dropped = (~is_num).sum()
    if n_dropped > 0:
        st.info(
            f"ℹ️ تم استبعاد **{n_dropped}** درس ما له رقم ترتيب بث حقيقي "
            f"(زي دروس \"إثراء\" بدون رقم) — ما يدخلون جدول البث."
        )
    return df[is_num].reset_index(drop=True)


# ============================================================

# ============================================================
# التقويم الدراسي — نحتاجه لتحديد الأيام المعطّلة تلقائيًا
# ============================================================
DAY_OFFSET = {"الأحد": 0, "الاثنين": 1, "الثلاثاء": 2, "الأربعاء": 3, "الخميس": 4}


def tf_is_day_holiday(semester: str, week: int, day: str,
                       calendar_weeks_df: pd.DataFrame, holidays_df: pd.DataFrame) -> bool:
    """يتأكد هل يوم معيّن (فصل + أسبوع + اسم يوم) يقع فعليًا ضمن فترة إجازة رسمية."""
    row = calendar_weeks_df[
        (calendar_weeks_df["الفصل_الدراسي"] == semester) & (calendar_weeks_df["الأسبوع"] == week)
    ]
    if row.empty:
        return False  # ما نقدر نتأكد من التاريخ، نتركه عادي بدل ما نخمّن
    start = pd.Timestamp(row.iloc[0]["تاريخ_البداية"])
    date = start + pd.Timedelta(days=DAY_OFFSET.get(day, 0))
    for _, h in holidays_df.iterrows():
        if pd.Timestamp(h["تاريخ_البداية"]) <= date <= pd.Timestamp(h["تاريخ_النهاية"]):
            return True
    return False


# ============================================================
# محرك "تعبئة قالب جدول بث جاهز" — تعطينه قالب أسبوع فاضي واحد بس،
# والبرنامج يستنسخه بعدد الأسابيع اللي يحتاجها فعليًا (حسب بيانات
# ملفات المواد المرفوعة)، ويحط "إجازة" تلقائيًا بالأيام المعطّلة
# حسب التقويم الرسمي، ويعبّي الباقي بمحتوى المواد.
# ============================================================
TF_DAYS_ORDER = ["الأحد", "الاثنين", "الثلاثاء", "الأربعاء", "الخميس"]
TF_PERIOD_NAME_TO_NUM = {
    "الاولي": 1, "الثانيه": 2, "الثالثه": 3, "الرابعه": 4,
    "الخامسه": 5, "السادسه": 6, "السابعه": 7, "الثامنه": 8,
}
TF_SKIP_SUBJECTS = {"اجازة"}  # خانات مو مواد حقيقية أصلًا داخل القالب (نادر، احتياط)


_TF_GRADE_ORDINALS = ["اول", "ثاني", "ثالث", "رابع", "خامس", "سادس", "سابع", "ثامن", "تاسع", "عاشر"]
_TF_GRADE_LEVELS = ["ابتدائي", "متوسط", "ثانوي", "تأهيلي", "تاهيلي"]


def tf_norm_grade(text: str) -> str:
    """يستخرج نمط الصف (الترتيب + المرحلة) من أي نص، بغضّ النظر عن أي كلمات
    زايدة قبله أو بعده (زي '1أول ابتدائي فكري' -> 'اول ابتدائي'). هذا مطلوب
    لأن قوالب بعض المسارات الخاصة (زي التربية الفكرية) تكتب اسم الصف بصياغة
    عامة ("أول ابتدائي") بينما ملفات موادها تضيف كلمة المسار ("فكري") لنفس
    الحقل — والمطابقة تصير بينهم فقط لأن كل مسار له قالبه وملفاته في جلسة
    توليد منفصلة عن المسارات الثانية، فما فيه خطر اختلاط."""
    text = _normalize_arabic(text)
    text = text.replace("اولي", "اول")  # "أولى" (صياغة مؤنّثة) تكافئ "أول"

    found_ordinal = next((o for o in _TF_GRADE_ORDINALS if o in text), None)
    found_level = next((l for l in _TF_GRADE_LEVELS if l in text), None)
    if found_ordinal and found_level:
        found_level = "تأهيلي" if found_level == "تاهيلي" else found_level
        return f"{found_ordinal} {found_level}"

    text = re.sub(r"^\d+", "", text).strip()
    text = re.sub(r"^ال", "", text).strip()
    return text


def tf_find_grade_blocks(ws):
    """يدور على كل بلوكات الصفوف داخل ورقة القالب (لازم يكون فيها 'اليوم' و'الحصة' برأس كل بلوك)."""
    blocks = []
    for r in range(1, ws.max_row + 1):
        a = _normalize_arabic(ws.cell(row=r, column=1).value or "")
        b = _normalize_arabic(ws.cell(row=r, column=2).value or "")
        if a == "اليوم" and b == "الحصه":
            grade_name = None
            for back in range(1, 5):
                v = ws.cell(row=r - back, column=3).value
                if v and str(v).strip():
                    grade_name = str(v).strip()
                    break
            blocks.append({"header_row": r, "grade": grade_name})
    return blocks


def tf_get_period_columns(ws, header_row):
    cols = {}
    c = 3
    while True:
        v = ws.cell(row=header_row, column=c).value
        if not v:
            break
        norm = _normalize_arabic(v)
        if norm in TF_PERIOD_NAME_TO_NUM:
            cols[c] = TF_PERIOD_NAME_TO_NUM[norm]
        c += 1
    return cols


def tf_get_slots_for_block(ws, block):
    header_row = block["header_row"]
    period_cols = tf_get_period_columns(ws, header_row)
    slots = []
    for day_idx, day in enumerate(TF_DAYS_ORDER):
        subject_row = header_row + 1 + day_idx * 3
        title_row = subject_row + 1
        link_row = subject_row + 2
        for col, period_num in period_cols.items():
            subject = ws.cell(row=subject_row, column=col).value
            if subject and str(subject).strip() and _normalize_arabic(subject) not in TF_SKIP_SUBJECTS:
                slots.append({
                    "day": day, "period": period_num,
                    "subject": str(subject).strip(),
                    "subject_cell": (subject_row, col),
                    "title_cell": (title_row, col),
                    "link_cell": (link_row, col),
                })
    return slots


# جديد: مجموعات مرادفات مخصّصة لمواد لها أسماء مختلفة تمامًا نصيًا
# لكنها نفس المادة فعليًا حسب تأكيد المستخدمة — المطابقة العادية
# (احتواء نصي) ما تكفي هنا لأن الأسماء لا تشترك بأي جزء نصي مشترك.
TF_SUBJECT_SYNONYM_GROUPS = [
    ["لغتي", "لغتي الخالدة", "لغتي الخالده", "اللغة العربية", "لغتي الجميلة"],
]


def _tf_in_same_synonym_group(a: str, b: str) -> bool:
    na = _tf_strip_al_per_word(_normalize_arabic(a))
    nb = _tf_strip_al_per_word(_normalize_arabic(b))
    for group in TF_SUBJECT_SYNONYM_GROUPS:
        norm_group = [_tf_strip_al_per_word(_normalize_arabic(g)) for g in group]
        a_in = any(na == g or na in g or g in na for g in norm_group)
        b_in = any(nb == g or nb in g or g in nb for g in norm_group)
        if a_in and b_in:
            return True
    return False


def _tf_strip_al_per_word(text: str) -> str:
    """يشيل "ال" التعريف و"و" العطف من بداية كل كلمة على حدة (مو النص كامل
    بس) — عشان 'تربية فنية' و'التربية الفنية'، وكذلك 'الحياتية و الاسرية'
    و'الحياتية والأسرية'، يتطابقوا حتى لو "و"/"ال" منفصلة بمسافة أو متصلة
    بالكلمة اللي بعدها، وبغضّ النظر عن ترتيب أو تكرار وجودها."""
    words = text.split()
    stripped = []
    for w in words:
        w = re.sub(r"^و", "", w)
        w = re.sub(r"^ال", "", w)
        if w:
            stripped.append(w)
    return "".join(stripped)


def _tf_subject_match(a: str, b: str) -> bool:
    """مطابقة مرنة بين اسمين لمادة — تتجاهل المسافات و"ال" التعريف على كل
    كلمة، وتقبل الاحتواء الجزئي (مثلاً 'رياضيات' يطابق 'الرياضيات').
    كمان تتحقق من مجموعات المرادفات المخصّصة (زي لغتي/اللغة العربية)."""
    na = _tf_strip_al_per_word(_normalize_arabic(a))
    nb = _tf_strip_al_per_word(_normalize_arabic(b))
    if not na or not nb:
        return False
    if na == nb or na in nb or nb in na:
        return True
    return _tf_in_same_synonym_group(a, b)


def _tf_normalize_subject_df(raw_df):
    """يرجّع {sheet_key: DataFrame مطبَّع} من DataFrame وحيد أو dict أوراق."""
    return {"_single": raw_df} if not isinstance(raw_df, dict) else raw_df


def _tf_is_numeric(v) -> bool:
    """فحص رقمي موثوق — يتعامل صح مع أرقام numpy (اللي فحص isinstance العادي يفشل معها بصمت)،
    وأيضًا يستبعد NaN صراحة (لأن float(nan) ينجح بدون خطأ رغم إنه مو رقم فعلي صالح)."""
    if v is None:
        return False
    if pd.isna(v):
        return False
    try:
        if isinstance(v, str):
            return False  # ما نعتبر النصوص أرقام حتى لو قابلة للتحويل، تفاديًا لتحويلات غير مقصودة
        float(v)
        return True
    except (ValueError, TypeError):
        return False


def tf_build_sequential_index(subject_dfs: dict):
    """
    جديد: بدل الفهرسة حسب الأسبوع، نبني قائمة مرتّبة لكل (مادة، صف) بنفس
    ترتيب الدروس الموجود في الملف (عمود ترتيب البث، وإلا ترتيب الصفوف نفسه) —
    وبعدين نسحب منها بالتتابع حسب عدد الخانات المطلوبة كل أسبوع، بغضّ النظر
    تمامًا عن أي رقم أسبوع مكتوب بملف المادة.

    subject_dfs: {اسم_المادة (من اسم الملف): DataFrame أو {sheet_name: DataFrame}}
    يرجّع:
    - index: {(subject_name_كما_رفعته, grade): [(عنوان، رابط), ...]} بترتيب البث الصحيح
    - summary: قائمة تشخيصية لعرضها للمستخدمة
    """
    index = {}
    summary = []
    for subject_name, raw_df in subject_dfs.items():
        for sheet_key, df in _tf_normalize_subject_df(raw_df).items():
            df = normalize_columns(df)
            if "عنوان_الدرس" not in df.columns:
                summary.append({
                    "الملف": subject_name, "الورقة": str(sheet_key),
                    "المشكلة": "ما لقينا عمود عنوان الدرس بعد المطابقة",
                })
                continue

            if "الصف" in df.columns and df["الصف"].notna().any():
                grade = tf_norm_grade(str(df["الصف"].dropna().iloc[0]))
            else:
                grade = tf_norm_grade(str(sheet_key))

            # جديد: اسم المادة يُقرأ من عمود "المادة" جوا الملف نفسه (بدل
            # الاعتماد على اسم الملف اللي رفعتيه) — أوثق، خصوصًا لو ملفات
            # نفس المادة تختلف تسميتها من ملف لآخر. لو العمود غير موجود أو
            # فاضي، نرجع لاسم الملف كخطة احتياطية.
            if "المادة" in df.columns and df["المادة"].notna().any():
                effective_subject = str(df["المادة"].dropna().iloc[0]).strip()
            else:
                effective_subject = subject_name

            # جديد: ما نرتّب حسب "ترتيب البث" إطلاقًا — نثق بترتيب الصفوف
            # الطبيعي بالملف كما هو، عشان أي درس (عادي أو إثرائي بدون رقم)
            # يبقى بمكانه الأصلي بالضبط، مو ينزاح لآخر القائمة.

            link_col = "رابط_يوتيوب" if "رابط_يوتيوب" in df.columns else None
            pairs = [
                (row["عنوان_الدرس"], row[link_col] if link_col else "")
                for _, row in df.iterrows()
            ]
            key = (effective_subject, grade)
            index.setdefault(key, []).extend(pairs)

            summary.append({
                "الملف": subject_name, "المادة_المكتشفة": effective_subject, "الصف_المكتشف": grade,
                "عدد_الدروس": len(pairs),
            })
    return index, summary


def tf_lookup_sequential_queue(index: dict, template_subject: str, grade: str):
    """يدور على أفضل تطابق مرن لاسم المادة داخل نفس الصف، ويرجّع مرجع القائمة
    الكاملة (مو نسخة) عشان المؤشر يتقدّم فيها بثبات عبر كل الأسابيع."""
    for (file_subject, file_grade), pairs in index.items():
        if file_grade == grade and _tf_subject_match(template_subject, file_subject):
            return pairs
    return None


def tf_generate_and_fill(template_wb, subject_dfs: dict, calendar_weeks_df=None, holidays_df=None,
                          target_semester: str = None):
    """
    template_wb: ملف فيه ورقة واحدة فاضية (قالب أسبوع عام يتكرر شكله كل أسبوع).
    يستنسخ هذي الورقة لعدد أسابيع الفصل الدراسي المختار (من ملف التقويم)، يحط
    "إجازة" تلقائيًا بالأيام المعطّلة رسميًا، ويعبّي الباقي بمحتوى المواد —
    **بالتتابع** (يأخذ من كل مادة عدد الدروس المطلوب لكل أسبوع، بنفس ترتيبها
    داخل ملفها، بغضّ النظر تمامًا عن أي رقم أسبوع مكتوب هناك).
    يرجّع: (الملف الناتج، تحذيرات، تشخيص المحتوى)
    """
    content_index, content_summary = tf_build_sequential_index(subject_dfs)
    warnings = []

    have_calendar = calendar_weeks_df is not None and holidays_df is not None
    if not have_calendar:
        warnings.append("⚠️ ملف التقويم مطلوب لتحديد عدد الأسابيع — ما فيه أسابيع تُولَّد بدونه.")
        return template_wb, warnings, content_summary

    semester = target_semester or "الأول"
    weeks_this_semester = sorted(
        calendar_weeks_df.loc[calendar_weeks_df["الفصل_الدراسي"] == semester, "الأسبوع"].dropna().unique()
    )
    weeks_this_semester = [int(w) for w in weeks_this_semester if _tf_is_numeric(w)]

    if not weeks_this_semester:
        warnings.append(f"⚠️ ما لقينا أي أسبوع للفصل {semester} في ملف التقويم.")
        return template_wb, warnings, content_summary

    blank_ws = template_wb.worksheets[0]
    blank_name = blank_ws.title
    blocks = tf_find_grade_blocks(blank_ws)  # الهيكل ثابت، نحسبه مرة وحدة من القالب الفاضي

    if not blocks:
        warnings.append("⚠️ ما لقينا أي بلوك صف داخل القالب (نبحث عن خلية \"اليوم\" وجنبها \"الحصة\") — تأكدي من شكل القالب.")
        return template_wb, warnings, content_summary

    # مؤشر تقدّم لكل (مادة، صف) — يتقدّم بثبات عبر كل الأسابيع بالتتابع
    cursors = {}

    for week in weeks_this_semester:
        new_ws = template_wb.copy_worksheet(blank_ws)
        new_ws.title = f"أسبوع {week}"[:31]

        holiday_days = set()
        for day in TF_DAYS_ORDER:
            if tf_is_day_holiday(semester, week, day, calendar_weeks_df, holidays_df):
                holiday_days.add(day)

        for block in blocks:
            grade = tf_norm_grade(block["grade"] or "")
            slots = tf_get_slots_for_block(new_ws, block)

            # نحط "إجازة" بالأيام المعطّلة كاملة، ونمسح خانات العنوان/الرابط فيها
            if holiday_days:
                period_cols = tf_get_period_columns(new_ws, block["header_row"])
                for day_idx, day in enumerate(TF_DAYS_ORDER):
                    if day not in holiday_days:
                        continue
                    subject_row = block["header_row"] + 1 + day_idx * 3
                    for col in period_cols:
                        new_ws.cell(row=subject_row, column=col).hyperlink = None
                        new_ws.cell(row=subject_row, column=col, value="إجازة")
                        for rr in (subject_row + 1, subject_row + 2):
                            c = new_ws.cell(row=rr, column=col)
                            c.hyperlink = None
                            c.value = None

            # جديد: نمسح كل خانات العنوان/الرابط (غير أيام الإجازة) أول شي،
            # قبل أي محاولة تعبئة — عشان لو القالب "الفاضي" فيه بقايا قديمة
            # (روابط من ملف سابق مثلاً)، ما تظهر وكأنها نتيجة صحيحة لما تفشل
            # مادة معيّنة تلقى محتوى لها.
            for s in slots:
                if s["day"] in holiday_days:
                    continue
                tr, tc = s["title_cell"]
                lr, lc = s["link_cell"]
                new_ws.cell(row=tr, column=tc).hyperlink = None
                new_ws.cell(row=tr, column=tc).value = None
                new_ws.cell(row=lr, column=lc).hyperlink = None
                new_ws.cell(row=lr, column=lc).value = None

            # جديد: نجمّع الخانات حسب هوية "المخزون" الفعلي اللي ترجعه المطابقة
            # (id(queue))، مو حسب نص المادة بالضبط — عشان "لغتي" و"اللغة العربية"
            # (نفس الملف، أسماء مختلفة) يتعاملوا كمجموعة وحدة، ونرتّبهم زمنيًا
            # (يوم ثم حصة) قبل السحب من المخزون، مو بترتيب ظهور نص المادة بالقالب.
            groups_by_queue = {}  # queue_key -> {"queue":..., "slots":[...], "subject_label":...}
            for s in slots:
                if s["day"] in holiday_days:
                    continue
                queue = tf_lookup_sequential_queue(content_index, s["subject"], grade)
                queue_key = (grade, id(queue)) if queue is not None else (grade, s["subject"], None)
                if queue_key not in groups_by_queue:
                    groups_by_queue[queue_key] = {"queue": queue, "slots": [], "subject_label": s["subject"]}
                groups_by_queue[queue_key]["slots"].append(s)

            for queue_key, group_data in groups_by_queue.items():
                queue = group_data["queue"]
                subject = group_data["subject_label"]
                # ترتيب زمني صحيح (يوم ثم حصة) بغضّ النظر عن ترتيب اكتشاف الخانات
                subject_slots = sorted(
                    group_data["slots"],
                    key=lambda s: (TF_DAYS_ORDER.index(s["day"]), s["period"]),
                )

                if queue_key not in cursors:
                    cursors[queue_key] = {"pos": 0}
                state = cursors[queue_key]

                if queue is None:
                    available = []
                else:
                    available = queue[state["pos"]: state["pos"] + len(subject_slots)]

                if len(available) < len(subject_slots):
                    warnings.append(
                        f"⚠️ \"{subject}\" — صف {block['grade']} — أسبوع {week} ({semester}): "
                        f"محتاجين {len(subject_slots)} درس، متوفر بس {len(available)} "
                        f"({'ما فيه ملف مرفوع لهذي المادة' if queue is None else 'خلص محتوى هذي المادة'})."
                    )

                for i, slot in enumerate(subject_slots):
                    if i >= len(available):
                        break
                    title, link = available[i]
                    if pd.isna(link) or str(link).strip() == "":
                        link = "لا يوجد رابط بث"
                    tr, tc = slot["title_cell"]
                    lr, lc = slot["link_cell"]
                    new_ws.cell(row=tr, column=tc).hyperlink = None
                    new_ws.cell(row=lr, column=lc).hyperlink = None
                    new_ws.cell(row=tr, column=tc, value=title)
                    new_ws.cell(row=lr, column=lc, value=link)

                if queue is not None:
                    state["pos"] += len(available)

    del template_wb[blank_name]
    return template_wb, warnings, content_summary


# ============================================================
# الواجهة
# ============================================================
# ============================================================
# جديد: محرك مقارنة جدولين (جدولنا مقابل جدول الفريق) — نفس القالب،
# نفحص كل خانة (مادة/عنوان درس/رابط) ونلوّن أي اختلاف بالأحمر في نسخة
# من ملف الفريق، ونرجّع قائمة بكل الاختلافات المكتشفة.
# ============================================================
_CMP_RED_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")


def _cmp_extract_week_num(sheet_name: str):
    """يستخرج رقم الأسبوع من اسم الورقة بغضّ النظر عن أي نص زايد حواليه،
    وبغضّ النظر هل الرقم مكتوب رقمًا (زي 'الأسبوع 3') أو كلمة عربية
    (زي 'الأسبوع الثالث') — عشان جدولنا وجدول الفريق يتقارنوا بنفس رقم
    الأسبوع حتى لو صياغة اسم الورقة مختلفة تمامًا."""
    m = re.search(r"(\d+)", sheet_name)
    if m:
        return int(m.group(1))

    norm = _normalize_arabic(sheet_name)
    for word, num in sorted(WEEK_WORD_TO_NUM.items(), key=lambda kv: -len(kv[0])):
        if _normalize_arabic(word) in norm:
            return num
    return None


def compare_schedules(our_wb, team_wb):
    """
    our_wb: الملف اللي طلّعه برنامجنا (المرجع الصحيح)
    team_wb: ملف الفريق (نلوّن الاختلافات فيه ونرجّعه)
    يرجّع: (team_wb الملوَّن، قائمة تفاصيل الاختلافات، قائمة تحذيرات هيكلية)

    المطابقة بين الأوراق تصير **برقم الأسبوع المستخرج من الاسم**، مو بتطابق
    نص اسم الورقة الكامل — عشان "تربية فكرية - الأسبوع 3" و"أسبوع 3"
    يتقارنوا صح حتى لو صياغة التسمية مختلفة بين الملفين.
    """
    diffs = []
    warnings = []

    our_by_week = {}
    for s in our_wb.sheetnames:
        w = _cmp_extract_week_num(s)
        if w is not None:
            our_by_week.setdefault(w, s)  # أول ورقة بنفس الرقم لو تكررت

    team_by_week = {}
    for s in team_wb.sheetnames:
        w = _cmp_extract_week_num(s)
        if w is not None:
            team_by_week.setdefault(w, s)

    common_weeks = sorted(set(our_by_week) & set(team_by_week))
    missing_in_team = sorted(set(our_by_week) - set(team_by_week))
    extra_in_team = sorted(set(team_by_week) - set(our_by_week))

    if missing_in_team:
        warnings.append(f"⚠️ أسابيع موجودة بجدولنا وناقصة بجدول الفريق: {', '.join(map(str, missing_in_team))}")
    if extra_in_team:
        warnings.append(f"⚠️ أسابيع زايدة بجدول الفريق مو موجودة بجدولنا: {', '.join(map(str, extra_in_team))}")

    for week in common_weeks:
        our_ws = our_wb[our_by_week[week]]
        team_ws = team_wb[team_by_week[week]]
        sheet_name = team_by_week[week]
        blocks = tf_find_grade_blocks(team_ws)

        for block in blocks:
            slots = tf_get_slots_for_block(team_ws, block)
            for slot in slots:
                tr, tc = slot["title_cell"]
                lr, lc = slot["link_cell"]

                for (r, c), field_name in [((tr, tc), "عنوان الدرس"), ((lr, lc), "رابط اليوتيوب")]:
                    our_val = our_ws.cell(row=r, column=c).value
                    team_val = team_ws.cell(row=r, column=c).value
                    our_norm = str(our_val).strip() if our_val is not None else ""
                    team_norm = str(team_val).strip() if team_val is not None else ""
                    if our_norm != team_norm:
                        team_ws.cell(row=r, column=c).fill = _CMP_RED_FILL
                        diffs.append({
                            "الورقة": sheet_name, "الصف": block["grade"], "اليوم": slot["day"],
                            "الحصة": slot["period"], "الحقل": field_name,
                            "عندنا": our_norm or "(فاضي)", "عند الفريق": team_norm or "(فاضي)",
                        })

    return team_wb, diffs, warnings


st.caption("النسخة الأولى (Prototype شخصي) — بُنيت للتعلّم والتجربة، بهوية التعليمية 2026")
st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)


app_mode = st.radio(
    "وش تبين تسوين؟",
    ["🧩 توليد جدول جديد", "🔍 مقارنة جدولين (جدولنا مقابل جدول الفريق)"],
    horizontal=True,
)
st.divider()

if app_mode == "🧩 توليد جدول جديد":
    st.markdown("""
    <div class="step-card">
        <div class="step-head">
            <div class="step-num">١</div>
            <p class="step-title">رفع الملفات</p>
        </div>
        <p class="step-sub">
            ارفعي ملف "جدول البث" — قالب أسبوع واحد فاضٍ (خانات المواد معبّاة، عنوان الدرس ورابط اليوتيوب فاضيين)،
            ملف أو أكثر لكل مادة (فيه عنوان الدرس ورابط اليوتيوب مرتبين بنفس تسلسل الحصص)،
            وملف التقويم الدراسي (فيه ورقتين: الأسابيع + الإجازات).
        </p>
    </div>
    """, unsafe_allow_html=True)

    col_a, col_b = st.columns(2)
    with col_a:
        template_file = st.file_uploader("📋 ملف جدول البث (القالب الفاضي — أسبوع واحد)", type=["xlsx"], key="tf_template")
        calendar_file = st.file_uploader(
            "🗓️ ملف التقويم الدراسي (الأسابيع + الإجازات)",
            type=["xlsx"], key="tf_calendar",
        )
    with col_b:
        subject_files = st.file_uploader(
            "📚 ملفات المواد (وحدة أو أكثر — اسم كل ملف يُعتبر اسم المادة)",
            type=["xlsx"], accept_multiple_files=True, key="tf_subjects",
        )
        target_semester = st.radio("📆 الفصل الدراسي المطلوب توليده", ["الأول", "الثاني"], horizontal=True)

    if template_file and subject_files and calendar_file:
        subject_dfs = {}
        for f in subject_files:
            subject_name = f.name.rsplit(".", 1)[0]
            subject_dfs[subject_name] = pd.read_excel(f, sheet_name=None)

        try:
            calendar_weeks_df = pd.read_excel(calendar_file, sheet_name="الأسابيع")
            holidays_df = pd.read_excel(calendar_file, sheet_name="الإجازات")
        except Exception as e:
            st.markdown(
                f'<div class="result-banner warn">⚠️ ما قدرنا نقرأ ملف التقويم — تأكدي إن فيه ورقتين بالاسم بالضبط '
                f'"الأسابيع" و"الإجازات" ({type(e).__name__}).</div>',
                unsafe_allow_html=True,
            )
            st.stop()

        st.markdown("""
        <div class="step-card">
            <div class="step-head">
                <div class="step-num">٢</div>
                <p class="step-title">توليد وتعبئة الجدول</p>
            </div>
            <p class="step-sub">اضغطي الزر وانتظري — البرنامج يولّد أسابيع الفصل المختار كاملة ويعبّيها من ملفات موادك.</p>
        </div>
        """, unsafe_allow_html=True)

        if st.button("🧩 ولّدي الأسابيع وعبّي القالب", type="primary"):
            with st.spinner("جاري التوليد والتعبئة..."):
                template_wb = openpyxl.load_workbook(template_file, data_only=True)
                filled_wb, tf_warnings, tf_summary = tf_generate_and_fill(
                    template_wb, subject_dfs, calendar_weeks_df, holidays_df, target_semester=target_semester
                )
                buf = io.BytesIO()
                filled_wb.save(buf)
            st.session_state["tf_filled_wb_bytes"] = buf.getvalue()
            st.session_state["tf_warnings"] = tf_warnings
            st.session_state["tf_summary"] = tf_summary
            st.session_state["tf_n_weeks"] = len(filled_wb.sheetnames)

        if st.session_state.get("tf_filled_wb_bytes"):
            tf_warnings = st.session_state.get("tf_warnings", [])
            tf_summary = st.session_state.get("tf_summary", [])
            n_warn = len(tf_warnings)
            n_weeks = st.session_state.get('tf_n_weeks', 0)

            st.markdown(f"""
            <div class="stat-row">
                <div class="stat-card ok"><div class="num">{n_weeks}</div><div class="lbl">أسبوع تم توليده</div></div>
                <div class="stat-card {'ok' if n_warn == 0 else 'warn'}"><div class="num">{n_warn}</div><div class="lbl">ملاحظة نقص محتوى</div></div>
                <div class="stat-card"><div class="num">{len(subject_files)}</div><div class="lbl">ملف مادة مرفوع</div></div>
            </div>
            """, unsafe_allow_html=True)

            with st.expander("🔍 تشخيص: وش اكتُشف من ملفات المواد اللي رفعتيها", expanded=(n_warn > 20)):
                if tf_summary:
                    st.table(pd.DataFrame(tf_summary))
                else:
                    st.write("ما انقرأ أي محتوى صالح من أي ملف مرفوع.")

            if n_warn == 0:
                st.markdown('<div class="result-banner ok">✅ تم تعبئة كل الخانات بدون أي نقص محتوى.</div>', unsafe_allow_html=True)
            else:
                st.markdown(
                    f'<div class="result-banner warn">⚠️ تم التوليد والتعبئة، لكن فيه <b>{n_warn}</b> ملاحظة '
                    f'(خانات ناقصة محتوى أو أسابيع غير معروفة):</div>',
                    unsafe_allow_html=True,
                )
                with st.expander("عرض كل الملاحظات", expanded=(n_warn <= 15)):
                    for w in tf_warnings:
                        st.write(w)

            st.download_button(
                label="⬇️ تحميل جدول البث الكامل (Excel)",
                data=st.session_state["tf_filled_wb_bytes"],
                file_name="جدول_البث_معبّى.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )
    else:
        st.markdown(
            '<div class="result-banner info">ℹ️ ارفعي ملف القالب الفاضي + ملف مادة واحد على الأقل + ملف التقويم الدراسي للمتابعة.</div>',
            unsafe_allow_html=True,
        )

else:
    st.markdown("""
    <div class="step-card">
        <div class="step-head">
            <div class="step-num">🔍</div>
            <p class="step-title">مقارنة جدولين</p>
        </div>
        <p class="step-sub">
            ارفعي جدول البث اللي طلّعه برنامجنا (المرجع)، وجدول البث اللي سوّاه الفريق يدويًا (بنفس القالب بالضبط) —
            وبنلوّن لك كل خانة مختلفة بينهم باللون الأحمر داخل نسخة من ملف الفريق.
        </p>
    </div>
    """, unsafe_allow_html=True)

    col_x, col_y = st.columns(2)
    with col_x:
        our_file = st.file_uploader("📄 جدول برنامجنا (المرجع)", type=["xlsx"], key="cmp_our")
    with col_y:
        team_file = st.file_uploader("👥 جدول الفريق (اللي بنقارنه)", type=["xlsx"], key="cmp_team")

    if our_file and team_file:
        if st.button("🔍 قارني الجدولين", type="primary"):
            with st.spinner("جاري المقارنة..."):
                our_wb = openpyxl.load_workbook(our_file, data_only=True)
                team_wb = openpyxl.load_workbook(team_file)
                colored_wb, diffs, cmp_warnings = compare_schedules(our_wb, team_wb)
                buf = io.BytesIO()
                colored_wb.save(buf)
            st.session_state["cmp_result_bytes"] = buf.getvalue()
            st.session_state["cmp_diffs"] = diffs
            st.session_state["cmp_warnings"] = cmp_warnings

        if st.session_state.get("cmp_result_bytes"):
            diffs = st.session_state.get("cmp_diffs", [])
            cmp_warnings = st.session_state.get("cmp_warnings", [])
            n_diff = len(diffs)

            st.markdown(f"""
            <div class="stat-row">
                <div class="stat-card {'ok' if n_diff == 0 else 'crit'}"><div class="num">{n_diff}</div><div class="lbl">خانة مختلفة</div></div>
            </div>
            """, unsafe_allow_html=True)

            for w in cmp_warnings:
                st.markdown(f'<div class="result-banner warn">{w}</div>', unsafe_allow_html=True)

            if n_diff == 0:
                st.markdown('<div class="result-banner ok">✅ الجدولين متطابقين تمامًا — ولا خانة مختلفة.</div>', unsafe_allow_html=True)
            else:
                st.markdown(
                    f'<div class="result-banner warn">⚠️ لقينا <b>{n_diff}</b> خانة مختلفة — مفصّلة تحت، ومُلوَّنة بالأحمر داخل الملف المُحمَّل.</div>',
                    unsafe_allow_html=True,
                )
                with st.expander("عرض كل الاختلافات", expanded=(n_diff <= 30)):
                    st.table(pd.DataFrame(diffs))

            st.download_button(
                label="⬇️ تحميل جدول الفريق مع تظليل الاختلافات (Excel)",
                data=st.session_state["cmp_result_bytes"],
                file_name="مقارنة_جدول_الفريق.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )
    else:
        st.markdown(
            '<div class="result-banner info">ℹ️ ارفعي الملفين (جدولنا وجدول الفريق) للمتابعة.</div>',
            unsafe_allow_html=True,
        )
