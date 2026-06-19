"""
DBATU Academic Result PDF → Structured Data Extractor
=====================================================
Converts DBATU examination result PDFs into structured, machine-readable
Excel, CSV, and JSON outputs.

Usage:
    python dbatu_extractor.py <pdf_file_or_folder> [--output-dir ./output] [--format xlsx,csv,json]

Examples:
    python dbatu_extractor.py result.pdf
    python dbatu_extractor.py result.pdf --output-dir ./output --format xlsx,csv,json
    python dbatu_extractor.py . --output-dir ./output
"""

import argparse
import json
import os
import re
import sys
from pathlib import Path

import pandas as pd
import pdfplumber


# =============================================================================
# Phase 2: Metadata Extraction
# =============================================================================

def extract_metadata(page1_lines: list[str]) -> dict:
    """
    Extract examination metadata from page 1 header.
    
    Expected:
        Line 0: "Dr. Babasaheb Ambedkar Technological University"
        Line 1: "EXAMINATION : {program} SEMESTER - {N} {session} {year} ( {type} ) Institute Record"
    """
    metadata = {
        "university": "",
        "program": "",
        "semester": 0,
        "session": "",
        "year": 0,
        "exam_type": "",
    }

    # Line 0: University name
    if page1_lines:
        metadata["university"] = page1_lines[0].strip()

    # Line 1: Examination details
    for line in page1_lines[:5]:
        m = re.search(
            r"EXAMINATION\s*:\s*(.+?)\s+SEMESTER\s*-\s*(\d+)\s+"
            r"(\w+)\s+(\d{4})\s*\(\s*(\w+)\s*\)\s*Institute\s*Record",
            line,
        )
        if m:
            metadata["program"] = m.group(1).strip()
            metadata["semester"] = int(m.group(2))
            metadata["session"] = m.group(3).strip()
            metadata["year"] = int(m.group(4))
            metadata["exam_type"] = m.group(5).strip()
            break

    return metadata


# =============================================================================
# Phase 3: Subject Auto-Discovery
# =============================================================================

def discover_subjects(page1_lines: list[str]) -> list[dict]:
    """
    Extract subject codes, names, and credits from page 1 subject description lines.
    
    These are the lines between the EXAMINATION line and the "Seat No." line,
    containing patterns like: CODE : SUBJECT NAME (Credit :N )
    Some subjects (GRADE-only) have no (Credit :N) suffix.
    
    Strategy: Split text on subject code boundaries, then parse each segment.
    """
    subjects = []

    # Find the range of subject description lines
    start_idx = None
    end_idx = None

    for i, line in enumerate(page1_lines):
        if "EXAMINATION" in line:
            start_idx = i + 1
        if start_idx is not None and "Seat No." in line:
            end_idx = i
            break

    if start_idx is None or end_idx is None:
        return subjects

    # Concatenate subject description lines
    subject_text = " ".join(page1_lines[start_idx:end_idx])

    # Subject codes look like: alphanumeric, >= 6 chars, start with letter or digit,
    # followed by " : "
    # Split text by finding all subject code boundaries
    # Pattern: look for CODE : NAME sequences
    code_pattern = re.compile(r"([A-Z0-9]{6,})\s*:\s*")
    
    # Find all code positions
    code_matches = list(code_pattern.finditer(subject_text))
    
    for idx, cm in enumerate(code_matches):
        code = cm.group(1)
        name_start = cm.end()
        
        # Name extends until the next subject code or end of text
        if idx + 1 < len(code_matches):
            name_end = code_matches[idx + 1].start()
        else:
            name_end = len(subject_text)
        
        name_text = subject_text[name_start:name_end].strip()
        
        # Check if this subject has (Credit :N )
        credit_match = re.search(r"\(Credit\s*:\s*(\d+)\s*\)\s*$", name_text)
        if credit_match:
            credit = int(credit_match.group(1))
            name = name_text[:credit_match.start()].strip()
        else:
            # GRADE-only subject — no credit specified
            credit = 0
            name = name_text.strip()
            # Clean up trailing junk
            name = re.sub(r"\s*\)\s*$", "", name)
        
        subjects.append({
            "code": code,
            "name": name,
            "credit": credit,
        })

    return subjects


def parse_marking_scheme(header_lines: list[str], subjects: list[dict]) -> dict:
    """
    Parse the column header block to extract marking scheme details.
    
    Returns a dict with:
        - subject_order: list of subject codes in column order
        - credits: list of credits per subject
        - ese_max: list of (max, min) tuples for ESE
        - ca_max: list of CA internal max marks
        - mid_max: list of MID internal max marks (may be shorter)
        - total_max: list of (max, min) tuples for total
        - grade_only_indices: set of indices that are GRADE-only
        - total_marks_max: int (e.g., 1050 or 700)
        - total_grade_points_max: float
        - total_credits_max: float
    """
    scheme = {
        "subject_order": [],
        "credits": [],
        "ese_max": [],
        "ca_max": [],
        "mid_max": [],
        "total_max": [],
        "grade_only_indices": set(),
        "total_marks_max": 0,
        "total_grade_points_max": 0.0,
        "total_credits_max": 0.0,
    }

    for line in header_lines:
        # Subject code row (may be truncated) — skip for now, use discovered subjects
        # CORE Tot.GrP.-250.00 Cr.25.00
        m = re.search(r"Tot\.GrP\.\-?([\d.]+)\s+Cr\.([\d.]+)", line)
        if m:
            scheme["total_grade_points_max"] = float(m.group(1))
            scheme["total_credits_max"] = float(m.group(2))

        # Total Marks(1050) or Total Marks(700)
        m = re.search(r"Total\s*Marks\((\d+)\)", line)
        if m:
            scheme["total_marks_max"] = int(m.group(1))

        # CREDIT row
        if line.strip().startswith("CREDIT"):
            credits = re.findall(r"\b(\d+)\b", line.split("CREDIT", 1)[1])
            scheme["credits"] = [int(c) for c in credits]

        # EXT TOTAL or ESE TOTAL row
        if re.match(r"\s*(EXT|ESE)\s+TOTAL", line):
            fractions = re.findall(r"(\d+/\d+)", line)
            scheme["ese_max"] = []
            for f in fractions:
                parts = f.split("/")
                scheme["ese_max"].append((int(parts[0]), int(parts[1])))

        # CA INTERNAL row
        if "CA INTERNAL" in line:
            nums = re.findall(r"\b(\d+)\b", line.split("CA INTERNAL", 1)[1])
            scheme["ca_max"] = [int(n) for n in nums]

        # MID INTERNAL row
        if "MID INTERNAL" in line:
            nums = re.findall(r"\b(\d+)\b", line.split("MID INTERNAL", 1)[1])
            scheme["mid_max"] = [int(n) for n in nums]

        # TOTAL row
        if re.match(r"\s*TOTAL\s+", line) and "TOTAL" in line and "Marks" not in line:
            tokens = line.split("TOTAL", 1)[1].strip().split()
            total_entries = []
            grade_indices = set()
            idx = 0
            for token in tokens:
                if "/" in token:
                    parts = token.split("/")
                    total_entries.append((int(parts[0]), int(parts[1])))
                    idx += 1
                elif token == "GRADE":
                    total_entries.append(("GRADE", "GRADE"))
                    grade_indices.add(idx)
                    idx += 1
            scheme["total_max"] = total_entries
            scheme["grade_only_indices"] = grade_indices

    return scheme


def build_subject_order(subjects: list[dict], scheme: dict) -> list[dict]:
    """
    Build the ordered list of subjects matching the column order.
    
    Assigns each subject its marking scheme data.
    """
    ordered = []
    num_subjects = len(subjects)

    for i, subj in enumerate(subjects):
        entry = dict(subj)  # copy
        entry["index"] = i

        # Credit
        if i < len(scheme["credits"]):
            entry["credit"] = scheme["credits"][i]

        # ESE — some subjects might not have ESE
        entry["has_ese"] = True
        if i < len(scheme.get("ese_max", [])):
            entry["ese_max"], entry["ese_min"] = scheme["ese_max"][i]
        else:
            entry["has_ese"] = False
            entry["ese_max"] = 0
            entry["ese_min"] = 0

        # Total
        if i < len(scheme.get("total_max", [])):
            t = scheme["total_max"][i]
            if t == ("GRADE", "GRADE"):
                entry["grade_only"] = True
                entry["total_max"] = 0
                entry["total_min"] = 0
            else:
                entry["grade_only"] = False
                entry["total_max"], entry["total_min"] = t
        else:
            entry["grade_only"] = False

        ordered.append(entry)

    return ordered


# =============================================================================
# Phase 4: Header/Footer Detection
# =============================================================================

HEADER_MARKERS = [
    "Dr. Babasaheb Ambedkar Technological University",
    "EXAMINATION :",
    "Seat No.",
    "SGPA",
    "Center Code",
    "CORE Tot.GrP.",
    "CREDIT",
    "EXT TOTAL",
    "ESE TOTAL",
    "CA INTERNAL",
    "MID INTERNAL",
    "INT TOTAL",
    "Total Marks(",
]

FOOTER_MARKERS = [
    "Cancel Seat No",
    "GRADE:",
    "Note :-",
    "AOO =",
    "Print By",
]


def is_header_line(line: str) -> bool:
    """Check if a line is part of the page header block."""
    stripped = line.strip()
    if not stripped:
        return False
    
    for marker in HEADER_MARKERS:
        if marker in stripped:
            return True
    
    # TOTAL line with fraction patterns (part of header)
    if re.match(r"^\s*TOTAL\s+(\d+/\d+|GRADE)", stripped):
        return True

    return False


def is_footer_line(line: str) -> bool:
    """Check if a line is part of the page footer block."""
    stripped = line.strip()
    for marker in FOOTER_MARKERS:
        if marker in stripped:
            return True
    return False


def extract_data_lines(page_text: str) -> list[str]:
    """
    Extract only student data lines from a page's text,
    removing header and footer lines.
    """
    lines = page_text.split("\n")
    data_lines = []
    in_data = False

    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue

        if is_footer_line(stripped):
            break  # stop at footer

        if not in_data:
            # Wait for TOTAL line (last header line) to pass
            if re.match(r"^\s*TOTAL\s+(\d+/\d+|GRADE)", stripped):
                in_data = True
                continue
            if is_header_line(stripped):
                continue
            # Check for continuation lines (truncated subject codes)
            # e.g., "1 8A 2 3 06 05D 0" — short alphanumeric, only in header zone
            if re.match(r"^[\dA-Z\s]+$", stripped) and len(stripped) < 40:
                continue
            # If we hit a student line without seeing TOTAL first,
            # we're already in data (can happen with unusual formats)
            if re.match(r"^[A-Z]?\d{13,16}\s+", stripped):
                in_data = True
                data_lines.append(stripped)
        else:
            # Once in data mode, include ALL lines (student marks, grades, etc.)
            # Only skip if it's clearly a repeated header on a new page
            # (this shouldn't happen since we break at footer)
            data_lines.append(stripped)

    return data_lines


# =============================================================================
# Phase 5: Student Record Extraction
# =============================================================================

# Seat number pattern: optional letter prefix + 13-16 digits
SEAT_PATTERN = re.compile(r"^([A-Z]?\d{13,16})\s+(.+?)\s+(\d{5})\s*-\s*(.+?)\s+(PASS|FAIL|ATKT|WITHHELD)\s*$")


def is_student_start(line: str) -> bool:
    """Check if a line starts a new student record."""
    return bool(SEAT_PATTERN.match(line.strip()))


def parse_student_info(line: str) -> dict | None:
    """Parse the student header line (seat no, name, institute, result)."""
    m = SEAT_PATTERN.match(line.strip())
    if not m:
        return None

    seat_no = m.group(1)
    raw_name = m.group(2).strip()
    institute_code = m.group(3)
    institute_name = m.group(4).strip()
    result = m.group(5)

    # Parse gender from (F) prefix
    gender = ""
    name = raw_name
    gm = re.match(r"^\(F\)\s*(.+)$", raw_name)
    if gm:
        gender = "F"
        name = gm.group(1).strip()

    return {
        "seat_no": seat_no,
        "name": name,
        "gender": gender,
        "institute_code": institute_code,
        "institute_name": institute_name,
        "result": result,
    }


def parse_ese_line(line: str, num_subjects_with_ese: int) -> tuple[list, float | None]:
    """
    Parse ESE marks line.
    Returns (ese_values, sgpa).
    
    Format: "034 048 040 032 037 039 035 039 028 036  7.74"
    Values can be 3-digit numbers, ZOO, or AB.
    SGPA is at the end (decimal number) — only present for PASS/some students.
    """
    stripped = line.strip()

    # Try to extract SGPA from end
    sgpa = None
    sgpa_match = re.search(r"\s+(\d+\.\d{2})\s*$", stripped)
    if sgpa_match:
        sgpa = float(sgpa_match.group(1))
        stripped = stripped[: sgpa_match.start()].strip()

    # Extract marks values
    values = re.findall(r"\b(\d{2,3}|ZOO|AB)\b", stripped)

    return values, sgpa


def parse_ca_line(line: str) -> tuple[str, str, list, float | None, float | None]:
    """
    Parse CA internal marks line.
    
    Format: "02533 (Whole) 017 041 020 018 020 015 018 018 019 053 054  193.50  25.00"
    Returns (center_code, center_type, ca_values, total_grade_points, total_credits)
    """
    stripped = line.strip()
    
    m = re.match(r"^(\d{5})\s+\((Whole|Part)\)\s+(.+)$", stripped)
    if not m:
        return "", "", [], None, None

    center_code = m.group(1)
    center_type = m.group(2)
    rest = m.group(3).strip()

    # Total grade points and credits are the last two decimal numbers
    total_gp = None
    total_cr = None
    gp_cr_match = re.search(r"\s+([\d.]+)\s+([\d.]+)\s*$", rest)
    if gp_cr_match:
        total_gp = float(gp_cr_match.group(1))
        total_cr = float(gp_cr_match.group(2))
        rest = rest[: gp_cr_match.start()].strip()

    # Extract CA values
    ca_values = re.findall(r"\b(\d{2,3}|ZOO|AB)\b", rest)

    return center_code, center_type, ca_values, total_gp, total_cr


def parse_total_marks_line(line: str) -> int | None:
    """Parse the total marks line (single number)."""
    stripped = line.strip()
    if re.match(r"^\d{2,4}$", stripped):
        return int(stripped)
    return None


def parse_mid_line(line: str) -> list:
    """
    Parse MID internal marks line.
    
    Format: "014 016 020 020 013 018 019 018"
    Can also contain ZOO.
    """
    values = re.findall(r"\b(\d{2,3}|ZOO|AB)\b", line.strip())
    return values


def parse_subject_totals_line(line: str) -> list:
    """
    Parse subject-wise total marks line.
    
    Format: "065 041 084 078 072 065 075 072 076 081 090"
    May contain "| |" for GRADE-only subjects.
    """
    stripped = line.strip()
    # Replace | | with a placeholder
    cleaned = re.sub(r"\|\s*\|", "GRADE_ONLY", stripped)
    
    parts = cleaned.split()
    values = []
    for p in parts:
        if p == "GRADE_ONLY":
            values.append("GRADE_ONLY")
        elif re.match(r"^\d{2,3}$", p):
            values.append(p)

    return values


def parse_grade_line(line: str) -> list:
    """
    Parse grade strings line.
    
    Format: "6.5/CD/19.5 8.5/AB/8.5 8/BB/24 8/BB/24 7.5/BC/15 ..."
    May contain "AU" for GRADE-only subjects.
    Grades may have (G-N) suffix for grace marks.
    """
    stripped = line.strip()

    # Find all grade patterns: value/letter/points or value/letter/points(G-N)
    grade_pattern = re.compile(
        r"(\d+\.?\d*)/([A-Z]{2})/([\d.]+)(?:\(G-(\d+)\))?"
    )

    results = []
    pos = 0
    tokens = stripped.split()
    
    i = 0
    while i < len(tokens):
        token = tokens[i]
        if token == "AU":
            # AU typically appears as "AU AU" — skip both
            results.append({
                "grade_value": None,
                "grade_letter": "AU",
                "grade_points": None,
                "grace_marks": None,
            })
            # Skip the next "AU" if present
            if i + 1 < len(tokens) and tokens[i + 1] == "AU":
                i += 2
            else:
                i += 1
            continue
        
        gm = grade_pattern.match(token)
        if gm:
            results.append({
                "grade_value": float(gm.group(1)),
                "grade_letter": gm.group(2),
                "grade_points": float(gm.group(3)),
                "grace_marks": int(gm.group(4)) if gm.group(4) else None,
            })
        i += 1

    return results


def safe_int(val: str) -> int | None:
    """Convert a mark value to int, handling ZOO and AB."""
    if val == "ZOO":
        return 0
    if val == "AB":
        return None
    try:
        return int(val)
    except (ValueError, TypeError):
        return None


def group_student_blocks(data_lines: list[str]) -> list[list[str]]:
    """
    Group data lines into blocks, where each block is one student's data.
    A new block starts at each student header line (seat number pattern).
    """
    blocks = []
    current_block = []

    for line in data_lines:
        if is_student_start(line):
            if current_block:
                blocks.append(current_block)
            current_block = [line]
        else:
            if current_block:  # only add to block if we've started one
                current_block.append(line)

    if current_block:
        blocks.append(current_block)

    return blocks


def parse_student_block(
    block: list[str],
    subjects: list[dict],
    scheme: dict,
    metadata: dict,
) -> dict | None:
    """
    Parse a student block (7-8 lines) into a structured record.
    
    Block structure:
        L0: Seat No + Name + Institute + Result
        L1: ESE marks + SGPA
        L2: Center Code + CA marks + Tot.GrP + Credits
        L3: Total marks
        L4: MID marks
        L5: [| |]  (Sem 7 only, for GRADE-only subjects)
        L5/L6: Subject totals
        L6/L7: Grade strings
    """
    if len(block) < 6:
        return None

    # L0: Student info
    info = parse_student_info(block[0])
    if not info:
        return None

    record = {
        "seat_no": info["seat_no"],
        "name": info["name"],
        "gender": info["gender"],
        "institute_code": info["institute_code"],
        "institute_name": info["institute_name"],
        "result": info["result"],
    }

    # L1: ESE marks + SGPA
    num_ese_subjects = sum(
        1 for s in subjects if not s.get("grade_only", False)
    )
    ese_values, sgpa = parse_ese_line(block[1], num_ese_subjects)
    record["sgpa"] = sgpa

    # L2: CA internal marks + Tot.GrP + Credits
    center_code, center_type, ca_values, total_gp, total_cr = parse_ca_line(block[2])
    record["center_code"] = center_code
    record["center_type"] = center_type
    record["total_grade_points"] = total_gp
    record["total_credits"] = total_cr

    # L3: Total marks
    total_marks = parse_total_marks_line(block[3])
    record["total_marks"] = total_marks

    # Determine which line contains what based on block length and content
    # For Sem 7 (with GRADE-only subjects): block has 8 lines with a "| |" line
    # For Sem 3 (all numeric): block has 7 lines
    
    mid_line_idx = 4
    
    # Find the grade line (last line should contain grade patterns like X/XX/X.X)
    grade_line_idx = None
    totals_line_idx = None
    
    for i in range(len(block) - 1, mid_line_idx, -1):
        if re.search(r"\d+\.?\d*/[A-Z]{2}/[\d.]+|AU", block[i]):
            grade_line_idx = i
            break
    
    if grade_line_idx is not None:
        totals_line_idx = grade_line_idx - 1
    
    # Fallback: if block has exactly the expected structure
    if grade_line_idx is None:
        if len(block) >= 7:
            grade_line_idx = len(block) - 1
            totals_line_idx = len(block) - 2

    # Parse MID marks
    mid_values = parse_mid_line(block[mid_line_idx]) if mid_line_idx < len(block) else []

    # Parse subject totals
    subject_totals = []
    if totals_line_idx is not None and totals_line_idx < len(block):
        subject_totals = parse_subject_totals_line(block[totals_line_idx])

    # Parse grades
    grades = []
    if grade_line_idx is not None and grade_line_idx < len(block):
        grades = parse_grade_line(block[grade_line_idx])

    # =========================================================================
    # Map marks to subjects
    # =========================================================================
    
    # Determine which subjects have ESE, which are GRADE-only
    grade_only_indices = scheme.get("grade_only_indices", set())
    
    ese_idx = 0
    ca_idx = 0
    mid_idx = 0
    total_idx = 0
    grade_idx = 0

    # Track which subjects have MID internal (from scheme)
    num_mid = len(scheme.get("mid_max", []))
    
    # Build mid_subject_map: for subjects that have MID internal
    # The MID INTERNAL row has fewer entries than total subjects
    # We need to figure out which subjects have MID
    # Strategy: subjects with MID are those that have "20" in MID INTERNAL row
    # This is positional — first N subjects that have MID get the MID values
    
    # For a simpler approach: subjects with ESE (not lab/grade-only) typically have MID
    # Labs (40/16 ESE) typically don't have MID
    # But let's use the count from scheme
    
    # Subjects with MID: first `num_mid` non-grade-only subjects that have mid
    # Actually, MID values are just positionally assigned to the first num_mid subjects
    # that are not lab subjects.
    # Looking at the data:
    #   Sem 3: 11 subjects, 8 MID values → subjects 0,2,3,4,5,6,7,8 (skipping 1 and 9,10)
    #   Subject 1 (VE308A, Credit 1, Total 50/20) — no MID
    #   Subjects 9,10 (Labs, 40/16) — no MID
    #   Sem 7: 9 subjects, 5 MID values → subjects 0,1,2,5,8 (skipping 3,4 labs, 6,7 grade-only)

    # Better approach: use ese_max to determine which subjects have MID
    # Subjects with ese_max of 60/20 AND NOT grade-only → have MID if total_max is 100/40
    # Subjects with ese_max of 40/16 → labs, no MID (CA is 60, no MID)
    # Subjects with total 50/20 → special (like VE308A), check if MID exists

    # Actually, simplest: subjects that have MID are those where:
    # total_max - ese_max - ca_max = mid_max (i.e., 100 - 60 - 20 = 20)
    # Labs: 100 - 40 - 60 = 0, no MID
    # VE308A: 50 - 60(???) — this doesn't add up.
    # Let me re-examine: VE308A has TOTAL 50/20, EXT TOTAL missing(?), CA 50
    # So VE308A = CA only (50 marks total), no ESE, no MID

    # Strategy: For each non-grade-only subject, check if it has MID
    # based on whether (total_max - ese_from_scheme - ca_from_scheme) > 0
    
    subjects_with_mid = []
    for i, subj in enumerate(subjects):
        if subj.get("grade_only", False):
            continue
        if i < len(scheme.get("ese_max", [])):
            ese_m = scheme["ese_max"][i][0]
        else:
            ese_m = 0
        if i < len(scheme.get("ca_max", [])):
            ca_m = scheme["ca_max"][i]
        else:
            ca_m = 0
        if i < len(scheme.get("total_max", [])):
            total = scheme["total_max"][i]
            if total == ("GRADE", "GRADE"):
                continue
            total_m = total[0]
        else:
            total_m = 0
        
        remaining = total_m - ese_m - ca_m
        if remaining > 0:
            subjects_with_mid.append(i)
    
    # Now map marks to each subject
    for i, subj in enumerate(subjects):
        code = subj["code"]
        is_grade_only = i in grade_only_indices or subj.get("grade_only", False)
        
        if is_grade_only:
            # Grade-only subject — no numeric marks
            record[f"{code}_ese"] = None
            record[f"{code}_ca"] = None
            record[f"{code}_mid"] = None
            record[f"{code}_total"] = None
            record[f"{code}_grade_letter"] = None
            record[f"{code}_grade_value"] = None
            record[f"{code}_grade_points"] = None
            record[f"{code}_grace_marks"] = None
            
            # Still consume grade entry
            if grade_idx < len(grades):
                g = grades[grade_idx]
                record[f"{code}_grade_letter"] = g["grade_letter"]
                record[f"{code}_grade_value"] = g["grade_value"]
                record[f"{code}_grade_points"] = g["grade_points"]
                record[f"{code}_grace_marks"] = g["grace_marks"]
                grade_idx += 1
            continue

        # ESE marks
        if ese_idx < len(ese_values):
            record[f"{code}_ese"] = safe_int(ese_values[ese_idx])
            ese_idx += 1
        else:
            record[f"{code}_ese"] = None

        # CA marks
        if ca_idx < len(ca_values):
            record[f"{code}_ca"] = safe_int(ca_values[ca_idx])
            ca_idx += 1
        else:
            record[f"{code}_ca"] = None

        # MID marks
        if i in subjects_with_mid:
            if mid_idx < len(mid_values):
                record[f"{code}_mid"] = safe_int(mid_values[mid_idx])
                mid_idx += 1
            else:
                record[f"{code}_mid"] = None
        else:
            record[f"{code}_mid"] = None

        # Subject total
        if total_idx < len(subject_totals):
            val = subject_totals[total_idx]
            if val == "GRADE_ONLY":
                record[f"{code}_total"] = None
            else:
                record[f"{code}_total"] = safe_int(val)
            total_idx += 1
        else:
            record[f"{code}_total"] = None

        # Grade
        if grade_idx < len(grades):
            g = grades[grade_idx]
            record[f"{code}_grade_letter"] = g["grade_letter"]
            record[f"{code}_grade_value"] = g["grade_value"]
            record[f"{code}_grade_points"] = g["grade_points"]
            record[f"{code}_grace_marks"] = g["grace_marks"]
            grade_idx += 1
        else:
            record[f"{code}_grade_letter"] = None
            record[f"{code}_grade_value"] = None
            record[f"{code}_grade_points"] = None
            record[f"{code}_grace_marks"] = None

    return record


# =============================================================================
# Phase 6 & 7: Process PDF → Export
# =============================================================================

def get_header_lines(page_text: str) -> list[str]:
    """Extract header lines from a page for scheme parsing."""
    lines = page_text.split("\n")
    header = []
    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue
        if is_header_line(stripped) or re.match(r"^\s*TOTAL\s+(\d+/\d+|GRADE)", stripped):
            header.append(stripped)
        elif re.match(r"^[A-Z]?\d{13,16}\s+", stripped):
            break  # hit first student
    return header


def process_pdf(pdf_path: str) -> dict:
    """
    Process a single DBATU result PDF.
    
    Returns:
        {
            "metadata": {...},
            "subjects": [...],
            "scheme": {...},
            "students": [...],
        }
    """
    print(f"\nProcessing: {pdf_path}")

    with pdfplumber.open(pdf_path) as pdf:
        num_pages = len(pdf.pages)
        print(f"  Pages: {num_pages}")

        # =====================================================================
        # Page 1: Extract metadata, subjects, and scheme
        # =====================================================================
        page1_text = pdf.pages[0].extract_text()
        page1_lines = page1_text.split("\n")

        metadata = extract_metadata(page1_lines)
        print(f"  Program: {metadata['program']}")
        print(f"  Semester: {metadata['semester']}, {metadata['session']} {metadata['year']}")

        subjects = discover_subjects(page1_lines)
        print(f"  Subjects discovered: {len(subjects)}")
        for s in subjects:
            print(f"    {s['code']}: {s['name']} (Cr: {s['credit']})")

        # Parse marking scheme from header
        header_lines = get_header_lines(page1_text)
        scheme = parse_marking_scheme(header_lines, subjects)
        print(f"  Total Marks Max: {scheme['total_marks_max']}")
        print(f"  Tot.GrP Max: {scheme['total_grade_points_max']}, Credits: {scheme['total_credits_max']}")
        print(f"  GRADE-only subject indices: {scheme['grade_only_indices']}")
        
        # Mark grade-only subjects
        for idx in scheme["grade_only_indices"]:
            if idx < len(subjects):
                subjects[idx]["grade_only"] = True
        
        # Ensure all subjects have grade_only flag
        for s in subjects:
            if "grade_only" not in s:
                s["grade_only"] = s.get("credit", 0) == 0

        # =====================================================================
        # All pages: Extract student records
        # =====================================================================
        all_students = []

        for page_no in range(num_pages):
            page = pdf.pages[page_no]
            page_text = page.extract_text()
            if not page_text:
                continue

            data_lines = extract_data_lines(page_text)
            blocks = group_student_blocks(data_lines)

            for block in blocks:
                student = parse_student_block(block, subjects, scheme, metadata)
                if student:
                    all_students.append(student)

        print(f"  Students extracted: {len(all_students)}")

    return {
        "metadata": metadata,
        "subjects": subjects,
        "scheme": scheme,
        "students": all_students,
    }


# =============================================================================
# Phase 7: Export Functions
# =============================================================================

def build_summary(students: list[dict], subjects: list[dict]) -> dict:
    """Build summary statistics from student data."""
    total = len(students)
    passed = sum(1 for s in students if s["result"] == "PASS")
    failed = sum(1 for s in students if s["result"] == "FAIL")
    atkt = sum(1 for s in students if s["result"] == "ATKT")

    sgpa_values = [s["sgpa"] for s in students if s["sgpa"] is not None]
    avg_sgpa = sum(sgpa_values) / len(sgpa_values) if sgpa_values else 0

    total_marks_values = [s["total_marks"] for s in students if s["total_marks"] is not None]
    avg_marks = sum(total_marks_values) / len(total_marks_values) if total_marks_values else 0

    # Subject-wise averages
    subject_stats = []
    for subj in subjects:
        code = subj["code"]
        if subj.get("grade_only", False):
            subject_stats.append({
                "subject_code": code,
                "subject_name": subj["name"],
                "type": "GRADE ONLY",
                "avg_total": None,
                "pass_count": None,
                "fail_count": None,
            })
            continue

        totals = [
            s.get(f"{code}_total")
            for s in students
            if s.get(f"{code}_total") is not None
        ]
        avg_total = sum(totals) / len(totals) if totals else 0

        # Count subject-level pass/fail based on grade
        pass_count = sum(
            1 for s in students
            if s.get(f"{code}_grade_letter") not in (None, "FF", "AU")
        )
        fail_count = sum(
            1 for s in students
            if s.get(f"{code}_grade_letter") == "FF"
        )

        subject_stats.append({
            "subject_code": code,
            "subject_name": subj["name"],
            "type": "REGULAR",
            "avg_total": round(avg_total, 2),
            "pass_count": pass_count,
            "fail_count": fail_count,
        })

    return {
        "total_students": total,
        "passed": passed,
        "failed": failed,
        "atkt": atkt,
        "pass_percentage": round(passed / total * 100, 2) if total > 0 else 0,
        "avg_sgpa": round(avg_sgpa, 2),
        "avg_total_marks": round(avg_marks, 2),
        "subject_stats": subject_stats,
    }


def export_xlsx(data: dict, output_path: str):
    """Export to Excel with multiple sheets."""
    metadata = data["metadata"]
    subjects = data["subjects"]
    students = data["students"]
    summary = build_summary(students, subjects)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        # Sheet 1: Metadata
        meta_df = pd.DataFrame([metadata])
        meta_df.to_excel(writer, sheet_name="Metadata", index=False)

        # Sheet 2: Subjects
        subj_df = pd.DataFrame(subjects)
        # Reorder columns
        cols = ["code", "name", "credit", "grade_only"]
        subj_df = subj_df[[c for c in cols if c in subj_df.columns]]
        subj_df.to_excel(writer, sheet_name="Subjects", index=False)

        # Sheet 3: Student Results
        if students:
            students_df = pd.DataFrame(students)
            # Reorder: student info first, then subject marks
            info_cols = [
                "seat_no", "name", "gender", "institute_code", "institute_name",
                "center_code", "center_type", "result", "sgpa",
                "total_grade_points", "total_credits", "total_marks",
            ]
            subj_cols = [c for c in students_df.columns if c not in info_cols]
            ordered_cols = [c for c in info_cols if c in students_df.columns] + sorted(subj_cols)
            students_df = students_df[ordered_cols]
            students_df.to_excel(writer, sheet_name="Student Results", index=False)

        # Sheet 4: Summary
        summary_info = {k: v for k, v in summary.items() if k != "subject_stats"}
        summary_df = pd.DataFrame([summary_info])
        summary_df.to_excel(writer, sheet_name="Summary", index=False)

        # Sheet 5: Subject Stats
        if summary["subject_stats"]:
            stats_df = pd.DataFrame(summary["subject_stats"])
            stats_df.to_excel(writer, sheet_name="Subject Stats", index=False)

    print(f"  Excel saved: {output_path}")


def export_csv(data: dict, output_dir: str, base_name: str):
    """Export to CSV files."""
    students = data["students"]
    subjects = data["subjects"]

    if students:
        students_df = pd.DataFrame(students)
        csv_path = os.path.join(output_dir, f"{base_name}_students.csv")
        students_df.to_csv(csv_path, index=False)
        print(f"  CSV saved: {csv_path}")

    subj_df = pd.DataFrame(subjects)
    csv_path = os.path.join(output_dir, f"{base_name}_subjects.csv")
    subj_df.to_csv(csv_path, index=False)
    print(f"  CSV saved: {csv_path}")


def export_json(data: dict, output_path: str):
    """Export to JSON."""
    # Convert for JSON serialization
    output = {
        "metadata": data["metadata"],
        "subjects": data["subjects"],
        "students": data["students"],
        "summary": build_summary(data["students"], data["subjects"]),
    }

    # Remove scheme from output (internal use only)
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(output, f, indent=2, ensure_ascii=False, default=str)

    print(f"  JSON saved: {output_path}")


# =============================================================================
# Main CLI
# =============================================================================

def make_base_name(metadata: dict) -> str:
    """Generate a base filename from metadata."""
    parts = []
    if metadata.get("program"):
        # Shorten program name
        prog = metadata["program"]
        if "Artificial Intelligence" in prog:
            parts.append("AI_DS")
        elif "Computer" in prog:
            parts.append("CSE")
        else:
            parts.append(prog[:20].replace(" ", "_"))
    parts.append(f"Sem{metadata.get('semester', 'X')}")
    if metadata.get("session"):
        parts.append(metadata["session"])
    if metadata.get("year"):
        parts.append(str(metadata["year"]))
    return "_".join(parts)


def main():
    parser = argparse.ArgumentParser(
        description="DBATU Academic Result PDF Extractor"
    )
    parser.add_argument(
        "input",
        help="Path to a PDF file or folder containing PDF files",
    )
    parser.add_argument(
        "--output-dir",
        default="./output",
        help="Output directory (default: ./output)",
    )
    parser.add_argument(
        "--format",
        default="xlsx",
        help="Output format(s), comma-separated: xlsx,csv,json (default: xlsx)",
    )
    args = parser.parse_args()

    input_path = Path(args.input)
    output_dir = Path(args.output_dir)
    formats = [f.strip().lower() for f in args.format.split(",")]

    # Collect PDF files
    if input_path.is_file() and input_path.suffix.lower() == ".pdf":
        pdf_files = [input_path]
    elif input_path.is_dir():
        pdf_files = sorted(input_path.glob("*.pdf"))
    else:
        print(f"Error: {input_path} is not a PDF file or directory")
        sys.exit(1)

    if not pdf_files:
        print("No PDF files found.")
        sys.exit(1)

    print(f"Found {len(pdf_files)} PDF file(s)")

    # Create output directory
    output_dir.mkdir(parents=True, exist_ok=True)

    # Process each PDF
    for pdf_file in pdf_files:
        try:
            data = process_pdf(str(pdf_file))
            base_name = make_base_name(data["metadata"])

            if "xlsx" in formats:
                xlsx_path = output_dir / f"{base_name}.xlsx"
                export_xlsx(data, str(xlsx_path))

            if "csv" in formats:
                export_csv(data, str(output_dir), base_name)

            if "json" in formats:
                json_path = output_dir / f"{base_name}.json"
                export_json(data, str(json_path))

        except Exception as e:
            print(f"\nError processing {pdf_file}: {e}")
            import traceback
            traceback.print_exc()

    print("\nDone!")


if __name__ == "__main__":
    main()
