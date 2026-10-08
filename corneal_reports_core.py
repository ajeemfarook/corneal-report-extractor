from pathlib import Path
import re
import pymupdf
import pandas as pd


# ============================================================
# CONFIGURATION
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

INPUT_DIR = BASE_DIR / "reports"
OUTPUT_DIR = BASE_DIR / "output"

OUTPUT_XLSX = OUTPUT_DIR / "corneal_reports.xlsx"
OUTPUT_CSV = OUTPUT_DIR / "corneal_reports.csv"
ERROR_CSV = OUTPUT_DIR / "corneal_reports_errors.csv"

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


# ============================================================
# GENERAL HELPERS
# ============================================================

def clean_text(value):
    if value is None:
        return None

    value = str(value)
    value = value.replace("\xa0", " ")
    value = re.sub(r"\s+", " ", value)

    return value.strip()


def safe_float(value):
    if value is None:
        return None

    try:
        value = str(value).replace(",", ".")
        return float(value)
    except Exception:
        return None


def safe_int(value):
    if value is None:
        return None

    try:
        return int(float(str(value).replace(",", ".")))
    except Exception:
        return None


def numeric_token(text):
    """
    Extract numeric value from a token.

    Examples:
        7.73       -> 7.73
        -1.90D     -> -1.90
        ax175°     -> 175
        0%         -> 0
    """

    if text is None:
        return None

    match = re.search(
        r"[-+]?\d+(?:[.,]\d+)?",
        str(text)
    )

    if not match:
        return None

    return safe_float(match.group(0))


def number_from_text(text):
    """
    Extract first numeric value from text.
    """

    if not text:
        return None

    match = re.search(
        r"[-+]?\d+(?:[.,]\d+)?",
        text
    )

    if not match:
        return None

    return safe_float(match.group(0))


def first_match(pattern, text, flags=re.IGNORECASE | re.MULTILINE):
    match = re.search(pattern, text, flags)

    if not match:
        return None

    return clean_text(match.group(1))


# ============================================================
# PDF WORD EXTRACTION
# ============================================================

def get_page_words(page):
    """
    Return PDF words with coordinates.

    Word structure:
        x0, y0, x1, y1, text, block, line, word
    """

    return page.get_text("words")


def words_in_region(
    words,
    x0,
    y0,
    x1,
    y1
):
    """
    Return words whose center is inside the specified region.
    """

    result = []

    for word in words:

        wx0, wy0, wx1, wy1, text = word[:5]

        center_x = (wx0 + wx1) / 2
        center_y = (wy0 + wy1) / 2

        if (
            x0 <= center_x <= x1
            and
            y0 <= center_y <= y1
        ):
            result.append(word)

    result.sort(
        key=lambda w: (
            round(w[1], 1),
            w[0]
        )
    )

    return result


def text_in_region(
    words,
    x0,
    y0,
    x1,
    y1
):
    region = words_in_region(
        words,
        x0,
        y0,
        x1,
        y1
    )

    return clean_text(
        " ".join(
            word[4]
            for word in region
        )
    )


def numbers_in_region(
    words,
    x0,
    y0,
    x1,
    y1,
    expected_count=None
):
    """
    Extract numeric values from a region.
    """

    region = words_in_region(
        words,
        x0,
        y0,
        x1,
        y1
    )

    values = []

    for word in region:

        value = numeric_token(word[4])

        if value is not None:
            values.append(value)

    if expected_count is not None:
        values = values[:expected_count]

    return values


# ============================================================
# LINE-BASED WORD PROCESSING
# ============================================================

def get_text_lines(words):
    """
    Group PDF words into visual text lines.

    This is more reliable than relying only on absolute
    coordinates.
    """

    lines = {}

    for word in words:

        x0, y0, x1, y1, text = word[:5]

        block = word[5] if len(word) > 5 else 0
        line = word[6] if len(word) > 6 else 0

        key = (block, line)

        if key not in lines:
            lines[key] = []

        lines[key].append(word)

    result = []

    for key, line_words in lines.items():

        line_words.sort(
            key=lambda w: w[0]
        )

        result.append(
            {
                "words": line_words,
                "x0": min(w[0] for w in line_words),
                "x1": max(w[2] for w in line_words),
                "y0": min(w[1] for w in line_words),
                "y1": max(w[3] for w in line_words),
                "text": clean_text(
                    " ".join(
                        w[4]
                        for w in line_words
                    )
                )
            }
        )

    result.sort(
        key=lambda line: (
            line["y0"],
            line["x0"]
        )
    )

    return result


# ============================================================
# PATIENT / REPORT INFORMATION
# ============================================================

def extract_patient_information(text, row):

    # --------------------------------------------------------
    # Patient
    # --------------------------------------------------------

    match = re.search(
        r"Patient\s+(.+?)\s+Gender",
        text,
        re.IGNORECASE | re.DOTALL
    )

    if match:

        patient = clean_text(
            match.group(1)
        )

        patient = patient.replace(
            "Patient",
            ""
        ).strip()

        row["Patient"] = patient

    # --------------------------------------------------------
    # Gender
    # --------------------------------------------------------

    match = re.search(
        r"Patient.*?Gender\s+([A-Za-z]+)",
        text,
        re.IGNORECASE | re.DOTALL
    )

    if match:
        row["Gender"] = clean_text(
            match.group(1)
        )

    # --------------------------------------------------------
    # Patient ID
    # --------------------------------------------------------

    match = re.search(
        r"\b(\d{3,})\s+Exam Date",
        text,
        re.IGNORECASE
    )

    if match:
        row["Patient_ID"] = match.group(1)

    # --------------------------------------------------------
    # Exam date / time
    # --------------------------------------------------------

    match = re.search(
        r"Exam Date\s+"
        r"(\d{2}/\d{2}/\d{4})\s+"
        r"(\d{2}:\d{2}:\d{2})",
        text,
        re.IGNORECASE
    )

    if match:

        row["Exam_Date"] = match.group(1)
        row["Exam_Time"] = match.group(2)

    # --------------------------------------------------------
    # Date of birth
    # --------------------------------------------------------

    match = re.search(
        r"Date\s+of\s+Birth\s+([^\n]+)",
        text,
        re.IGNORECASE
    )

    if match:

        dob = clean_text(
            match.group(1)
        )

        if dob and dob not in [
            "Date of Birth",
            "****"
        ]:
            row["Date_of_Birth"] = dob


# ============================================================
# SIM-K
# ============================================================

def extract_sim_k(words, text, row):
    """
    Extract Sim-K K1, K2 and CYL for OD and OS.

    Expected PDF layout:

        Sim-K
        K1              K2              CYL
        7.91 @ 166°    7.77 @ 76°      -0.74D ax166°
        7.95 @ 173°    7.72 @ 83°      -1.27D ax173°

    Output fields:

        OD_SimK_K1_D
        OD_SimK_K1_axis_deg
        OD_SimK_K2_D
        OD_SimK_K2_axis_deg
        OD_SimK_CYL_D
        OD_SimK_CYL_axis_deg

        OS_SimK_K1_D
        OS_SimK_K1_axis_deg
        OS_SimK_K2_D
        OS_SimK_K2_axis_deg
        OS_SimK_CYL_D
        OS_SimK_CYL_axis_deg

    The K1/K2 values are powers in diopters (D), not millimeters.

    The important fix is that extraction, validation, debug output and
    Excel column ordering all use the same *_D field names.
    """

    # ------------------------------------------------------------
    # Initialize all fields so every row has the same schema.
    # ------------------------------------------------------------
    fields = [
        "OD_SimK_K1_D",
        "OD_SimK_K1_axis_deg",
        "OD_SimK_K2_D",
        "OD_SimK_K2_axis_deg",
        "OD_SimK_CYL_D",
        "OD_SimK_CYL_axis_deg",
        "OS_SimK_K1_D",
        "OS_SimK_K1_axis_deg",
        "OS_SimK_K2_D",
        "OS_SimK_K2_axis_deg",
        "OS_SimK_CYL_D",
        "OS_SimK_CYL_axis_deg",
    ]

    for field in fields:
        row.setdefault(field, None)

    # ------------------------------------------------------------
    # Helper: parse K1/K2 value such as:
    #     7.91 @ 166°
    # ------------------------------------------------------------
    def parse_k_cell(cell_text):
        if not cell_text:
            return None, None

        match = re.search(
            r"([-+]?\d+(?:[.,]\d+)?)\s*@\s*"
            r"(\d+(?:[.,]\d+)?)\s*°?",
            cell_text,
            re.IGNORECASE,
        )

        if match:
            return (
                safe_float(match.group(1)),
                safe_float(match.group(2)),
            )

        # Fallback if the PDF separates @ / axis unexpectedly.
        numbers = re.findall(
            r"[-+]?\d+(?:[.,]\d+)?",
            cell_text,
        )

        if len(numbers) >= 2:
            return (
                safe_float(numbers[0]),
                safe_float(numbers[1]),
            )

        return None, None

    # ------------------------------------------------------------
    # Helper: parse CYL cell such as:
    #     -0.74D ax166°
    # ------------------------------------------------------------
    def parse_cyl_cell(cell_text):
        if not cell_text:
            return None, None

        match = re.search(
            r"([-+]?\d+(?:[.,]\d+)?)\s*D?\s*"
            r"(?:ax|@)\s*"
            r"(\d+(?:[.,]\d+)?)\s*°?",
            cell_text,
            re.IGNORECASE,
        )

        if match:
            return (
                safe_float(match.group(1)),
                safe_float(match.group(2)),
            )

        numbers = re.findall(
            r"[-+]?\d+(?:[.,]\d+)?",
            cell_text,
        )

        if len(numbers) >= 2:
            return (
                safe_float(numbers[0]),
                safe_float(numbers[1]),
            )

        return None, None

    # ------------------------------------------------------------
    # Coordinate-based extraction.
    #
    # For this report template the header positions are approximately:
    #
    # OD K1  x=52
    # OD K2  x=143
    # OD CYL x=229
    # OS K1  x=347
    # OS K2  x=438
    # OS CYL x=524
    #
    # We do NOT use these exact values as the only method. Instead,
    # we first locate the K1/K2/CYL header words and then use their
    # actual x positions from the PDF.
    # ------------------------------------------------------------
    lines = get_text_lines(words)

    simk_line = None
    header_line = None

    for line in lines:
        if re.search(r"\bSim[- ]?K\b", line["text"], re.IGNORECASE):
            simk_line = line
            break

    if simk_line is not None:
        for line in lines:
            if line["y0"] < simk_line["y1"]:
                continue

            # Header is immediately below Sim-K.
            if line["y0"] > simk_line["y1"] + 30:
                break

            upper = line["text"].upper()

            if (
                re.search(r"\bK1\b", upper)
                and re.search(r"\bK2\b", upper)
                and re.search(r"\bCYL\b", upper)
            ):
                header_line = line
                break

    # ------------------------------------------------------------
    # Locate the six header centers.
    # ------------------------------------------------------------
    header_x = {}
    header_y1_values = []

    # The PDF places K1/K2/CYL as separate PDF text lines, even though
    # they appear on the same visual row. Therefore do not depend on
    # get_text_lines() finding one combined header line.
    if simk_line is not None:
        for word in words:
            x0, y0, x1, y1, word_text = word[:5]
            label = clean_text(word_text).upper()

            if label not in {"K1", "K2", "CYL"}:
                continue

            if not (simk_line["y1"] <= y0 <= simk_line["y1"] + 30):
                continue

            center_x = (x0 + x1) / 2.0
            eye = "OD" if center_x < 300 else "OS"
            key = f"{eye}_{label}"
            header_x[key] = center_x
            header_y1_values.append(y1)

    # ------------------------------------------------------------
    # Fixed-position fallback for the known report template.
    # This is only used when the PDF does not expose the header
    # words in a usable way.
    # ------------------------------------------------------------
    default_header_x = {
        "OD_K1": 60.0,
        "OD_K2": 150.0,
        "OD_CYL": 241.0,
        "OS_K1": 354.0,
        "OS_K2": 445.0,
        "OS_CYL": 535.0,
    }

    for key, value in default_header_x.items():
        header_x.setdefault(key, value)

    # ------------------------------------------------------------
    # Determine the y position of the value row.
    # Normally it is immediately below the header line.
    # ------------------------------------------------------------
    if header_y1_values:
        # Start just below the visual header row. This prevents header
        # labels such as K1/K2/CYL from being included in the value cell.
        header_bottom = max(header_y1_values)
        value_y_min = header_bottom + 0.5
        value_y_max = header_bottom + 30
    elif header_line is not None:
        value_y_min = header_line["y1"] + 0.5
        value_y_max = header_line["y1"] + 30
    elif simk_line is not None:
        value_y_min = simk_line["y1"] + 14
        value_y_max = simk_line["y1"] + 50
    else:
        # Known template fallback.
        value_y_min = 430
        value_y_max = 450

    # ------------------------------------------------------------
    # Collect words from the Sim-K value row.
    # ------------------------------------------------------------
    value_words = []

    for word in words:
        x0, y0, x1, y1, word_text = word[:5]
        center_y = (y0 + y1) / 2.0

        if value_y_min <= center_y <= value_y_max:
            value_words.append(word)

    # ------------------------------------------------------------
    # Assign words to the nearest Sim-K header.
    # ------------------------------------------------------------
    cells = {
        "OD_K1": [],
        "OD_K2": [],
        "OD_CYL": [],
        "OS_K1": [],
        "OS_K2": [],
        "OS_CYL": [],
    }

    for word in value_words:
        x0, y0, x1, y1, word_text = word[:5]
        center_x = (x0 + x1) / 2.0

        nearest_key = None
        nearest_distance = float("inf")

        for key, header_center in header_x.items():
            distance = abs(center_x - header_center)

            if distance < nearest_distance:
                nearest_distance = distance
                nearest_key = key

        # Prevent unrelated words from being pulled into Sim-K.
        if nearest_key is not None and nearest_distance <= 55:
            cells[nearest_key].append(word)

    # ------------------------------------------------------------
    # Convert cells to text.
    # ------------------------------------------------------------
    cell_text = {}

    for key, cell_words in cells.items():
        cell_words.sort(key=lambda w: w[0])
        cell_text[key] = clean_text(
            " ".join(w[4] for w in cell_words)
        )

    # ------------------------------------------------------------
    # Parse OD.
    # ------------------------------------------------------------
    od_k1, od_k1_axis = parse_k_cell(cell_text["OD_K1"])
    od_k2, od_k2_axis = parse_k_cell(cell_text["OD_K2"])
    od_cyl, od_cyl_axis = parse_cyl_cell(cell_text["OD_CYL"])

    row["OD_SimK_K1_D"] = od_k1
    row["OD_SimK_K1_axis_deg"] = od_k1_axis
    row["OD_SimK_K2_D"] = od_k2
    row["OD_SimK_K2_axis_deg"] = od_k2_axis
    row["OD_SimK_CYL_D"] = od_cyl
    row["OD_SimK_CYL_axis_deg"] = od_cyl_axis

    # ------------------------------------------------------------
    # Parse OS.
    # ------------------------------------------------------------
    os_k1, os_k1_axis = parse_k_cell(cell_text["OS_K1"])
    os_k2, os_k2_axis = parse_k_cell(cell_text["OS_K2"])
    os_cyl, os_cyl_axis = parse_cyl_cell(cell_text["OS_CYL"])

    row["OS_SimK_K1_D"] = os_k1
    row["OS_SimK_K1_axis_deg"] = os_k1_axis
    row["OS_SimK_K2_D"] = os_k2
    row["OS_SimK_K2_axis_deg"] = os_k2_axis
    row["OS_SimK_CYL_D"] = os_cyl
    row["OS_SimK_CYL_axis_deg"] = os_cyl_axis

    # ------------------------------------------------------------
    # Text fallback.
    #
    # If coordinate extraction failed, find the Sim-K block and
    # parse six values in the expected order.
    # ------------------------------------------------------------
    required_k_fields = [
        "OD_SimK_K1_D",
        "OD_SimK_K1_axis_deg",
        "OD_SimK_K2_D",
        "OD_SimK_K2_axis_deg",
        "OS_SimK_K1_D",
        "OS_SimK_K1_axis_deg",
        "OS_SimK_K2_D",
        "OS_SimK_K2_axis_deg",
    ]

    if any(row.get(field) is None for field in required_k_fields):
        simk_text = ""

        if simk_line is not None:
            simk_y0 = simk_line["y0"]

            # Collect nearby text lines only.
            nearby = []
            for line in lines:
                if simk_y0 <= line["y0"] <= simk_y0 + 55:
                    nearby.append(line["text"])

            simk_text = clean_text(" ".join(nearby)) or ""

        if not simk_text:
            # Last fallback: use the text surrounding the literal Sim-K.
            match = re.search(
                r"Sim[- ]?K(.{0,300})",
                text,
                re.IGNORECASE | re.DOTALL,
            )
            if match:
                simk_text = clean_text(match.group(0)) or ""

        # Extract the four K values in order:
        # OD K1, OD K2, OS K1, OS K2.
        k_matches = re.findall(
            r"([-+]?\d+(?:[.,]\d+)?)\s*@\s*"
            r"(\d+(?:[.,]\d+)?)\s*°?",
            simk_text,
            re.IGNORECASE,
        )

        if len(k_matches) >= 4:
            fallback_values = [
                (safe_float(a), safe_float(b))
                for a, b in k_matches[:4]
            ]

            od_k1, od_k1_axis = fallback_values[0]
            od_k2, od_k2_axis = fallback_values[1]
            os_k1, os_k1_axis = fallback_values[2]
            os_k2, os_k2_axis = fallback_values[3]

            if row.get("OD_SimK_K1_D") is None:
                row["OD_SimK_K1_D"] = od_k1
            if row.get("OD_SimK_K1_axis_deg") is None:
                row["OD_SimK_K1_axis_deg"] = od_k1_axis

            if row.get("OD_SimK_K2_D") is None:
                row["OD_SimK_K2_D"] = od_k2
            if row.get("OD_SimK_K2_axis_deg") is None:
                row["OD_SimK_K2_axis_deg"] = od_k2_axis

            if row.get("OS_SimK_K1_D") is None:
                row["OS_SimK_K1_D"] = os_k1
            if row.get("OS_SimK_K1_axis_deg") is None:
                row["OS_SimK_K1_axis_deg"] = os_k1_axis

            if row.get("OS_SimK_K2_D") is None:
                row["OS_SimK_K2_D"] = os_k2
            if row.get("OS_SimK_K2_axis_deg") is None:
                row["OS_SimK_K2_axis_deg"] = os_k2_axis

        # CYL fallback.
        cyl_matches = re.findall(
            r"([-+]?\d+(?:[.,]\d+)?)\s*D\s*"
            r"(?:ax|@)\s*"
            r"(\d+(?:[.,]\d+)?)\s*°?",
            simk_text,
            re.IGNORECASE,
        )

        if len(cyl_matches) >= 2:
            od_cyl, od_cyl_axis = cyl_matches[0]
            os_cyl, os_cyl_axis = cyl_matches[1]

            if row.get("OD_SimK_CYL_D") is None:
                row["OD_SimK_CYL_D"] = safe_float(od_cyl)
            if row.get("OD_SimK_CYL_axis_deg") is None:
                row["OD_SimK_CYL_axis_deg"] = safe_float(od_cyl_axis)

            if row.get("OS_SimK_CYL_D") is None:
                row["OS_SimK_CYL_D"] = safe_float(os_cyl)
            if row.get("OS_SimK_CYL_axis_deg") is None:
                row["OS_SimK_CYL_axis_deg"] = safe_float(os_cyl_axis)



# ============================================================
# CORNEA DECENTRATION
# ============================================================

def extract_cornea_decentration(words, row):

    # OD rho
    values = numbers_in_region(
        words,
        150,
        460,
        230,
        480
    )

    if values:
        row[
            "OD_Cornea_Decentration_rho_mm"
        ] = values[0]

    # OD theta
    values = numbers_in_region(
        words,
        230,
        460,
        300,
        480
    )

    if values:
        row[
            "OD_Cornea_Decentration_theta_deg"
        ] = values[0]

    # OS rho
    values = numbers_in_region(
        words,
        440,
        460,
        510,
        480
    )

    if values:
        row[
            "OS_Cornea_Decentration_rho_mm"
        ] = values[0]

    # OS theta
    values = numbers_in_region(
        words,
        510,
        460,
        610,
        480
    )

    if values:
        row[
            "OS_Cornea_Decentration_theta_deg"
        ] = values[0]


# ============================================================
# CORNEA DIAMETER
# ============================================================

def extract_cornea_diameter(words, row):

    values = numbers_in_region(
        words,
        180,
        475,
        270,
        500
    )

    if values:
        row[
            "OD_Cornea_Diameter_mm"
        ] = values[0]

    values = numbers_in_region(
        words,
        480,
        475,
        570,
        500
    )

    if values:
        row[
            "OS_Cornea_Diameter_mm"
        ] = values[0]


# ============================================================
# PHOTOPIC PUPIL DECENTRATION
# ============================================================

def extract_pupil_decentration(words, row):

    # OD X
    values = numbers_in_region(
        words,
        150,
        490,
        220,
        515
    )

    if values:
        row[
            "OD_Photopic_Pupil_X_mm"
        ] = values[0]

    # OD Y
    values = numbers_in_region(
        words,
        220,
        490,
        300,
        515
    )

    if values:
        row[
            "OD_Photopic_Pupil_Y_mm"
        ] = values[0]

    # OS X
    values = numbers_in_region(
        words,
        445,
        490,
        520,
        515
    )

    if values:
        row[
            "OS_Photopic_Pupil_X_mm"
        ] = values[0]

    # OS Y
    values = numbers_in_region(
        words,
        515,
        490,
        610,
        515
    )

    if values:
        row[
            "OS_Photopic_Pupil_Y_mm"
        ] = values[0]


# ============================================================
# PHOTOPIC PUPIL DIAMETER
# ============================================================

def extract_pupil_diameter(words, row):

    values = numbers_in_region(
        words,
        190,
        505,
        260,
        530
    )

    if values:
        row[
            "OD_Photopic_Pupil_Diameter_mm"
        ] = values[0]

    values = numbers_in_region(
        words,
        485,
        505,
        560,
        530
    )

    if values:
        row[
            "OS_Photopic_Pupil_Diameter_mm"
        ] = values[0]


# ============================================================
# AVERAGE PUPILLAR POWER
# ============================================================

def extract_pupillar_power(words, text, row):
    """
    Extract:

        4.5mm: 44.56 D
        3mm:   44.63 D

    IMPORTANT:
    We do NOT simply take every number in the region.

    The old method would see:

        4.5
        44.56
        3
        44.63

    and could incorrectly save 4.5 as the power.

    This function specifically looks for the value AFTER
    4.5mm and AFTER 3mm.
    """

    # ========================================================
    # First: use text regex
    # ========================================================

    patterns_4_5 = [
        r"4\.5\s*mm\s*[:\-]?\s*"
        r"([-+]?\d+(?:[.,]\d+)?)\s*D",

        r"4\.5\s*mm\s*[:\-]?\s*"
        r"([-+]?\d+(?:[.,]\d+)?)",

        r"4\.5\s*mm.*?"
        r"([-+]?\d+(?:[.,]\d+)?)\s*D",
    ]

    patterns_3 = [
        r"3\s*mm\s*[:\-]?\s*"
        r"([-+]?\d+(?:[.,]\d+)?)\s*D",

        r"3\s*mm\s*[:\-]?\s*"
        r"([-+]?\d+(?:[.,]\d+)?)",

        r"3\s*mm.*?"
        r"([-+]?\d+(?:[.,]\d+)?)\s*D",
    ]

    # ========================================================
    # Split left and right eye areas using words
    # ========================================================

    lines = get_text_lines(words)

    power_lines = []

    for line in lines:

        if re.search(
            r"4\.5\s*mm",
            line["text"],
            re.IGNORECASE
        ) and re.search(
            r"3\s*mm",
            line["text"],
            re.IGNORECASE
        ):

            power_lines.append(line)

    # ========================================================
    # Process each eye independently
    # ========================================================

    for eye, xmin, xmax in [
        ("OD", 0, 310),
        ("OS", 310, 650)
    ]:

        eye_text_parts = []

        for line in power_lines:

            # Determine whether line belongs to eye
            center_x = (
                line["x0"] + line["x1"]
            ) / 2

            if xmin <= center_x <= xmax:

                eye_text_parts.append(
                    line["text"]
                )

        # If line grouping does not work, use region text
        if not eye_text_parts:

            region_text = text_in_region(
                words,
                xmin,
                510,
                xmax,
                555
            )

            if region_text:
                eye_text_parts.append(
                    region_text
                )

        eye_text = clean_text(
            " ".join(eye_text_parts)
        )

        if not eye_text:
            continue

        # ----------------------------------------------------
        # 4.5 mm power
        # ----------------------------------------------------

        power_4_5 = None

        for pattern in patterns_4_5:

            match = re.search(
                pattern,
                eye_text,
                re.IGNORECASE
            )

            if match:

                power_4_5 = safe_float(
                    match.group(1)
                )

                break

        # ----------------------------------------------------
        # 3 mm power
        # ----------------------------------------------------

        power_3 = None

        for pattern in patterns_3:

            match = re.search(
                pattern,
                eye_text,
                re.IGNORECASE
            )

            if match:

                power_3 = safe_float(
                    match.group(1)
                )

                break

        if power_4_5 is not None:

            row[
                f"{eye}_Avg_Pupillar_Power_4_5mm_D"
            ] = power_4_5

        if power_3 is not None:

            row[
                f"{eye}_Avg_Pupillar_Power_3mm_D"
            ] = power_3

    # ========================================================
    # Strong fallback:
    # Extract from entire text using eye order
    # ========================================================

    all_power_matches = re.findall(
        r"""
        4\.5\s*mm\s*[:\-]?\s*
        ([-+]?\d+(?:[.,]\d+)?)\s*D?
        .*?
        3\s*mm\s*[:\-]?\s*
        ([-+]?\d+(?:[.,]\d+)?)\s*D?
        """,
        text,
        re.IGNORECASE | re.VERBOSE
    )

    if len(all_power_matches) >= 1:

        m = all_power_matches[0]

        if row.get(
            "OD_Avg_Pupillar_Power_4_5mm_D"
        ) is None:

            row[
                "OD_Avg_Pupillar_Power_4_5mm_D"
            ] = safe_float(m[0])

        if row.get(
            "OD_Avg_Pupillar_Power_3mm_D"
        ) is None:

            row[
                "OD_Avg_Pupillar_Power_3mm_D"
            ] = safe_float(m[1])

    if len(all_power_matches) >= 2:

        m = all_power_matches[1]

        if row.get(
            "OS_Avg_Pupillar_Power_4_5mm_D"
        ) is None:

            row[
                "OS_Avg_Pupillar_Power_4_5mm_D"
            ] = safe_float(m[0])

        if row.get(
            "OS_Avg_Pupillar_Power_3mm_D"
        ) is None:

            row[
                "OS_Avg_Pupillar_Power_3mm_D"
            ] = safe_float(m[1])


# ============================================================
# AK
# ============================================================

def extract_ak(words, row):

    # OD
    values = numbers_in_region(
        words,
        20,
        550,
        300,
        590
    )

    if len(values) >= 4:

        row["OD_AK_D"] = values[0]
        row["OD_AK_AGC_D_per_mm"] = values[1]
        row["OD_AK_SI_D"] = values[2]
        row["OD_AK_Kpi_percent"] = values[3]

    # OS
    values = numbers_in_region(
        words,
        315,
        550,
        620,
        590
    )

    if len(values) >= 4:

        row["OS_AK_D"] = values[0]
        row["OS_AK_AGC_D_per_mm"] = values[1]
        row["OS_AK_SI_D"] = values[2]
        row["OS_AK_Kpi_percent"] = values[3]


# ============================================================
# KERATOCONUS COMPATIBILITY
# ============================================================

def extract_keratoconus(words, row):

    od_text = text_in_region(
        words,
        20,
        585,
        300,
        630
    )

    os_text = text_in_region(
        words,
        315,
        585,
        620,
        630
    )

    if od_text:

        row[
            "OD_Topography_Keratoconus_Assessment"
        ] = od_text

    if os_text:

        row[
            "OS_Topography_Keratoconus_Assessment"
        ] = os_text


# ============================================================
# KERATOREFRACTIVE INDEXES
# ============================================================

def extract_keratorefractive_indexes(words, row):

    # OD
    values = numbers_in_region(
        words,
        20,
        675,
        300,
        710
    )

    if len(values) >= 4:

        row["OD_SD_D"] = values[0]
        row["OD_SAI_D"] = values[1]
        row["OD_Asphericity_e"] = values[2]
        row["OD_Kc_D"] = values[3]

    # OS
    values = numbers_in_region(
        words,
        315,
        675,
        620,
        710
    )

    if len(values) >= 4:

        row["OS_SD_D"] = values[0]
        row["OS_SAI_D"] = values[1]
        row["OS_Asphericity_e"] = values[2]
        row["OS_Kc_D"] = values[3]


# ============================================================
# AXIAL LENGTH
# ============================================================

def extract_axial_length(words, row):

    # OD
    values = numbers_in_region(
        words,
        100,
        720,
        220,
        755
    )

    if values:

        row[
            "OD_Axial_Length_mm"
        ] = values[0]

    # OS
    values = numbers_in_region(
        words,
        400,
        720,
        520,
        755
    )

    if values:

        row[
            "OS_Axial_Length_mm"
        ] = values[0]


# ============================================================
# DEVICE INFORMATION
# ============================================================

def extract_device_information(text, row):

    # --------------------------------------------------------
    # Device / software
    # --------------------------------------------------------

    match = re.search(
        r"(MYOPIA\s+EXPERT\s+700\s*-\s*E\s+V[\d.]+)",
        text,
        re.IGNORECASE
    )

    if match:

        row["Device_Software"] = clean_text(
            match.group(1)
        )

    # --------------------------------------------------------
    # Serial number
    # --------------------------------------------------------

    match = re.search(
        r"SN:\s*(\d+)",
        text,
        re.IGNORECASE
    )

    if match:

        row[
            "Device_Serial_Number"
        ] = match.group(1)

    # --------------------------------------------------------
    # Report generated
    # --------------------------------------------------------

    matches = re.findall(
        r"\b\d{2}/\d{2}/\d{4}\s+"
        r"\d{2}:\d{2}:\d{2}\b",
        text
    )

    if len(matches) >= 2:

        row[
            "Report_Generated"
        ] = matches[-1]


# ============================================================
# RAW TEXT
# ============================================================

def extract_raw_text(doc):

    pages = []

    for page_number, page in enumerate(
        doc,
        start=1
    ):

        text = page.get_text(
            "text"
        )

        pages.append(
            f"--- PAGE {page_number} ---\n"
            f"{text}"
        )

    return "\n".join(pages)


# ============================================================
# DEBUG EXTRACTION
# ============================================================

def print_debug_information(
    pdf_path,
    row
):

    print()
    print("  Extracted important values:")

    fields = [

        # Sim-K
        "OD_SimK_K1_D",
        "OD_SimK_K1_axis_deg",
        "OD_SimK_K2_D",
        "OD_SimK_K2_axis_deg",
        "OD_SimK_CYL_D",
        "OD_SimK_CYL_axis_deg",

        "OS_SimK_K1_D",
        "OS_SimK_K1_axis_deg",
        "OS_SimK_K2_D",
        "OS_SimK_K2_axis_deg",
        "OS_SimK_CYL_D",
        "OS_SimK_CYL_axis_deg",

        # Pupillar
        "OD_Avg_Pupillar_Power_4_5mm_D",
        "OD_Avg_Pupillar_Power_3mm_D",

        "OS_Avg_Pupillar_Power_4_5mm_D",
        "OS_Avg_Pupillar_Power_3mm_D",
    ]

    for field in fields:

        print(
            f"    {field}: "
            f"{row.get(field)}"
        )


# ============================================================
# PROCESS ONE PDF
# ============================================================

def process_pdf(pdf_path):

    row = {
        "Source_File": pdf_path.name,
        "Source_Path": str(pdf_path),
        "Status": "OK",
    }

    errors = []

    doc = None

    try:

        doc = pymupdf.open(
            pdf_path
        )

        if len(doc) == 0:

            raise ValueError(
                "PDF contains no pages"
            )

        # ----------------------------------------------------
        # First page
        # ----------------------------------------------------

        page = doc[0]

        words = get_page_words(
            page
        )

        text = page.get_text(
            "text"
        )

        if not text.strip():

            raise ValueError(
                "No selectable text found. "
                "This PDF may be image/scanned based."
            )

        # ----------------------------------------------------
        # Patient
        # ----------------------------------------------------

        extract_patient_information(
            text,
            row
        )

        # ----------------------------------------------------
        # Sim-K
        # ----------------------------------------------------

        extract_sim_k(
            words,
            text,
            row
        )

        # ----------------------------------------------------
        # Cornea data
        # ----------------------------------------------------

        extract_cornea_decentration(
            words,
            row
        )

        extract_cornea_diameter(
            words,
            row
        )

        extract_pupil_decentration(
            words,
            row
        )

        extract_pupil_diameter(
            words,
            row
        )

        # ----------------------------------------------------
        # IMPORTANT:
        # Average Pupillar Power
        # ----------------------------------------------------

        extract_pupillar_power(
            words,
            text,
            row
        )

        # ----------------------------------------------------
        # AK
        # ----------------------------------------------------

        extract_ak(
            words,
            row
        )

        # ----------------------------------------------------
        # Keratoconus
        # ----------------------------------------------------

        extract_keratoconus(
            words,
            row
        )

        # ----------------------------------------------------
        # Keratorefractive indexes
        # ----------------------------------------------------

        extract_keratorefractive_indexes(
            words,
            row
        )

        # ----------------------------------------------------
        # Axial length
        # ----------------------------------------------------

        extract_axial_length(
            words,
            row
        )

        # ----------------------------------------------------
        # Device
        # ----------------------------------------------------

        extract_device_information(
            text,
            row
        )

        # ----------------------------------------------------
        # Raw text
        # ----------------------------------------------------

        row["Raw_Text"] = extract_raw_text(
            doc
        )

        # ----------------------------------------------------
        # Validation
        # ----------------------------------------------------

        important_fields = [

            "Patient_ID",
            "Exam_Date",

            "OD_SimK_K1_D",
            "OD_SimK_K2_D",
            "OD_SimK_CYL_D",
            "OD_SimK_CYL_axis_deg",

            "OS_SimK_K1_D",
            "OS_SimK_K2_D",
            "OS_SimK_CYL_D",
            "OS_SimK_CYL_axis_deg",

            "OD_Avg_Pupillar_Power_4_5mm_D",
            "OD_Avg_Pupillar_Power_3mm_D",

            "OS_Avg_Pupillar_Power_4_5mm_D",
            "OS_Avg_Pupillar_Power_3mm_D",

            "OD_Axial_Length_mm",
            "OS_Axial_Length_mm",
        ]

        missing = []

        for field in important_fields:

            if row.get(field) is None:

                missing.append(field)

        if missing:

            row["Status"] = "WARNING"

            errors.append(
                "Missing fields: "
                + ", ".join(missing)
            )

        # ----------------------------------------------------
        # Debug output
        # ----------------------------------------------------

        print_debug_information(
            pdf_path,
            row
        )

    except Exception as e:

        row["Status"] = "ERROR"

        errors.append(
            f"{type(e).__name__}: "
            f"{str(e)}"
        )

    finally:

        if doc is not None:

            doc.close()

    row["Error_Message"] = " | ".join(
        errors
    )

    return row


# ============================================================
# PROCESS ALL PDFS
# ============================================================

def main():

    print()
    print("=" * 80)
    print("CORNEAL REPORT PDF → EXCEL / CSV")
    print("=" * 80)
    print()

    # --------------------------------------------------------
    # Check input
    # --------------------------------------------------------

    if not INPUT_DIR.exists():

        print(
            f"ERROR: Input folder does not exist:\n"
            f"{INPUT_DIR}"
        )

        return

    # --------------------------------------------------------
    # Find PDFs
    # --------------------------------------------------------

    pdf_files = sorted(
        INPUT_DIR.glob("*.pdf")
    )

    if not pdf_files:

        print(
            f"No PDF files found in:\n"
            f"{INPUT_DIR}"
        )

        return

    print(
        f"Found {len(pdf_files)} PDF files."
    )

    print()

    rows = []

    # --------------------------------------------------------
    # Process PDFs
    # --------------------------------------------------------

    for index, pdf_path in enumerate(
        pdf_files,
        start=1
    ):

        print(
            f"[{index}/{len(pdf_files)}] "
            f"{pdf_path.name}"
        )

        row = process_pdf(
            pdf_path
        )

        rows.append(row)

    # ========================================================
    # CREATE DATAFRAME
    # ========================================================

    df = pd.DataFrame(
        rows
    )

    # ========================================================
    # COLUMN ORDER
    # ========================================================

    first_columns = [

        "Source_File",
        "Source_Path",
        "Status",
        "Error_Message",

        # Patient
        "Patient_ID",
        "Patient",
        "Gender",
        "Date_of_Birth",
        "Exam_Date",
        "Exam_Time",

        # ----------------------------
        # OD Sim-K
        # ----------------------------

        "OD_SimK_K1_D",
        "OD_SimK_K1_axis_deg",

        "OD_SimK_K2_D",
        "OD_SimK_K2_axis_deg",

        "OD_SimK_CYL_D",
        "OD_SimK_CYL_axis_deg",

        # ----------------------------
        # OS Sim-K
        # ----------------------------

        "OS_SimK_K1_D",
        "OS_SimK_K1_axis_deg",

        "OS_SimK_K2_D",
        "OS_SimK_K2_axis_deg",

        "OS_SimK_CYL_D",
        "OS_SimK_CYL_axis_deg",

        # ----------------------------
        # OD Cornea
        # ----------------------------

        "OD_Cornea_Decentration_rho_mm",
        "OD_Cornea_Decentration_theta_deg",

        "OD_Cornea_Diameter_mm",

        "OD_Photopic_Pupil_X_mm",
        "OD_Photopic_Pupil_Y_mm",

        "OD_Photopic_Pupil_Diameter_mm",

        "OD_Avg_Pupillar_Power_4_5mm_D",
        "OD_Avg_Pupillar_Power_3mm_D",

        # ----------------------------
        # OS Cornea
        # ----------------------------

        "OS_Cornea_Decentration_rho_mm",
        "OS_Cornea_Decentration_theta_deg",

        "OS_Cornea_Diameter_mm",

        "OS_Photopic_Pupil_X_mm",
        "OS_Photopic_Pupil_Y_mm",

        "OS_Photopic_Pupil_Diameter_mm",

        "OS_Avg_Pupillar_Power_4_5mm_D",
        "OS_Avg_Pupillar_Power_3mm_D",

        # ----------------------------
        # OD AK
        # ----------------------------

        "OD_AK_D",
        "OD_AK_AGC_D_per_mm",
        "OD_AK_SI_D",
        "OD_AK_Kpi_percent",

        # ----------------------------
        # OS AK
        # ----------------------------

        "OS_AK_D",
        "OS_AK_AGC_D_per_mm",
        "OS_AK_SI_D",
        "OS_AK_Kpi_percent",

        # ----------------------------
        # Keratoconus
        # ----------------------------

        "OD_Topography_Keratoconus_Assessment",
        "OS_Topography_Keratoconus_Assessment",

        # ----------------------------
        # Keratorefractive
        # ----------------------------

        "OD_SD_D",
        "OD_SAI_D",
        "OD_Asphericity_e",
        "OD_Kc_D",

        "OS_SD_D",
        "OS_SAI_D",
        "OS_Asphericity_e",
        "OS_Kc_D",

        # ----------------------------
        # Axial length
        # ----------------------------

        "OD_Axial_Length_mm",
        "OS_Axial_Length_mm",

        # ----------------------------
        # Device
        # ----------------------------

        "Device_Software",
        "Device_Serial_Number",
        "Report_Generated",

        # ----------------------------
        # Raw
        # ----------------------------

        "Raw_Text",
    ]

    # Keep only columns that exist
    existing_first_columns = [
        c
        for c in first_columns
        if c in df.columns
    ]

    remaining_columns = [
        c
        for c in df.columns
        if c not in existing_first_columns
    ]

    df = df[
        existing_first_columns
        + remaining_columns
    ]

    # ========================================================
    # SAVE CSV
    # ========================================================

    df_without_raw = df.drop(
        columns=["Raw_Text"],
        errors="ignore"
    )

    df_without_raw.to_csv(
        OUTPUT_CSV,
        index=False,
        encoding="utf-8-sig"
    )

    # ========================================================
    # ERROR REPORT
    # ========================================================

    error_df = df[
        df["Status"].isin(
            [
                "ERROR",
                "WARNING"
            ]
        )
    ].copy()

    error_df.to_csv(
        ERROR_CSV,
        index=False,
        encoding="utf-8-sig"
    )

    # ========================================================
    # SAVE EXCEL
    # ========================================================

    with pd.ExcelWriter(
        OUTPUT_XLSX,
        engine="openpyxl"
    ) as writer:

        # ----------------------------------------------------
        # Main data
        # ----------------------------------------------------

        df_without_raw.to_excel(
            writer,
            sheet_name="Extracted_Data",
            index=False
        )

        # ----------------------------------------------------
        # Raw text
        # ----------------------------------------------------

        raw_df = df[
            [
                "Source_File",
                "Patient_ID",
                "Raw_Text"
            ]
        ].copy()

        raw_df.to_excel(
            writer,
            sheet_name="Raw_Text",
            index=False
        )

        # ----------------------------------------------------
        # Errors
        # ----------------------------------------------------

        error_df.to_excel(
            writer,
            sheet_name="Errors",
            index=False
        )

    # ========================================================
    # SUMMARY
    # ========================================================

    successful = len(
        df[
            df["Status"] == "OK"
        ]
    )

    warnings = len(
        df[
            df["Status"] == "WARNING"
        ]
    )

    errors = len(
        df[
            df["Status"] == "ERROR"
        ]
    )

    print()
    print("=" * 80)
    print("FINISHED")
    print("=" * 80)

    print(
        f"Total reports : {len(df)}"
    )

    print(
        f"Successful    : {successful}"
    )

    print(
        f"Warnings      : {warnings}"
    )

    print(
        f"Errors        : {errors}"
    )

    print()

    print(
        f"Excel:\n{OUTPUT_XLSX}"
    )

    print(
        f"CSV:\n{OUTPUT_CSV}"
    )

    print(
        f"Errors:\n{ERROR_CSV}"
    )

    print()
    print("=" * 80)


# ============================================================
# RUN
# ============================================================

if __name__ == "__main__":
    main()