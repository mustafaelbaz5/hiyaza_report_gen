import re


def clean_text(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def normalized_key(value) -> str:
    text = clean_text(value)
    text = re.sub(r"\s+", " ", text)
    text = text.replace("ـ", "")
    return text.strip()


def safe_filename(value: str) -> str:
    text = normalized_key(value)
    text = re.sub(r'[<>:"/\\|?*]', "_", text)
    text = re.sub(r"\s+", "_", text)
    text = re.sub(r"_+", "_", text)
    return text.strip(" ._") or "تقرير_تسليم"


def tidy_association_name(value: str) -> str:
    text = normalized_key(value)
    text = re.sub(r"\s*-\s*", " - ", text)
    return text


def infer_association_type(name: str, sheet_name: str = "") -> str:
    combined = f"{name} {sheet_name}"
    if "إصلاح" in combined or "اصلاح" in combined:
        return "إصلاح زراعي"
    if "ائتمان" in combined:
        return "ائتمان زراعي"
    return "ائتمان زراعي"


def infer_village_name(association_name: str) -> str:
    text = tidy_association_name(association_name)
    for token in (" - الائتمان الزراعي", " - الإصلاح الزراعي", " - الاصلاح الزراعي",
                  " للائتمان الزراعي", " للإصلاح الزراعي"):
        text = text.replace(token, "")
    text = re.sub(r"^جمعية\s+", "", text).strip()
    return text or tidy_association_name(association_name)

