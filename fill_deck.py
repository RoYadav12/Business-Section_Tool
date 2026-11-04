#!/usr/bin/env python3
"""
fill_deck.py — Copy research data from a RefractOne-style Excel workbook
(sheets: "Overview", "IT Vendor", "Business Themes") into the RefractOne
Business Demo PPTX template, replacing the "xxxx" placeholders.

Usage:
    python3 fill_deck.py --excel COMPANY.xlsm --pptx Sample.pptx \
        --company "Welspun Enterprises" --output filled.pptx

Design notes
------------
- This is built for the *specific* Sample.pptx template layout (shape
  positions were mapped by hand). Re-run it for a new company by swapping
  in a new Excel workbook that follows the same three-sheet schema — you
  do not need to touch the code, just point --excel at the new file.
- Wherever the Excel data does not contain a value the template asks for
  (e.g. Ownership, Headquarter, Net Income, Market Cap, some timeline
  years, Key Developments table), the script writes "N/A" or a clearly
  flagged placeholder instead of guessing — see FLAG below. Search the
  output deck for FLAG to find everything that still needs a human to
  fill in.
- Editing preserves the formatting of the first run in every text box
  (per the pptx skill: assigning text_frame.text collapses formatting,
  so we always assign to a specific run instead).
"""

import argparse
import copy
import re
from datetime import datetime
import openpyxl
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from pptx.util import Pt

FLAG = "xxxx"


# --------------------------------------------------------------------------
# 1. Excel parsing
# --------------------------------------------------------------------------

def _find_sheet(wb, prefix):
    """Sheet names can get silently truncated/renamed by Excel (31-char
    limit), so match by prefix instead of an exact name."""
    for name in wb.sheetnames:
        if name.startswith(prefix):
            return wb[name]
    return None


def parse_excel(path):
    wb = openpyxl.load_workbook(path, data_only=True, keep_vba=True)
    ov = wb["Overview"]
    rows = [[c.value for c in row] for row in ov.iter_rows()]

    data = {
        "business_description": None,
        "segments": {},          # name -> description
        "segment_order": [],
        "timeline": {},          # int year -> text
        "peers": [],             # list of (name, revenue, employees)
        "it_vendors": [],
        "business_themes": [],
        "tech_themes": [],
        "esg_themes": [],        # 'Sustainability & ESG Themes' sheet
        "challenges": [],        # 'Challenges & Growth Opportunities' sheet
        "challenge_headers": ("Key Challenges", "Opportunities"),
        "recent_developments": [],
    }

    section = None
    current_segment = None
    current_peer = None
    i = 0
    while i < len(rows):
        row = rows[i]
        a, b, c = (row + [None, None, None])[:3]
        if a == "Business Description":
            data["business_description"] = b
        elif a == "Business Segments":
            section = "segments"
        elif a == "Business Timeline":
            section = "timeline"
        elif a == "Key Triggers & Recent Developments":
            section = "triggers"
        elif a == "Peer Companies":
            section = "peers"
        elif section == "segments" and a and b:
            data["segments"][a] = b
            data["segment_order"].append(a)
        elif section == "timeline" and isinstance(a, int):
            data["timeline"][a] = b
        elif section == "peers" and isinstance(a, str) and a.strip():
            if a.startswith("Revenue"):
                if current_peer:
                    current_peer["revenue"] = a.replace("Revenue", "").strip()
            elif a.startswith("Employees"):
                if current_peer:
                    current_peer["employees"] = a.replace("Employees", "").strip()
                    data["peers"].append(current_peer)
                    current_peer = None
            else:
                current_peer = {"name": a}
        i += 1

    # IT Vendor sheet
    itv = wb["IT Vendor"]
    it_rows = list(itv.iter_rows(values_only=True))
    header, *body = it_rows
    for r in body:
        if not r or not r[0]:
            continue
        rec = dict(zip(header, r))
        data["it_vendors"].append(rec)

    # Every theme list below ends up with the same normalised shape so the
    # slide fillers don't care which workbook layout it came from:
    #   headline / summary (one-line intro) / points (bullets) / impact_points
    # (the legacy supporting_instances / commentary / strategic_impact keys
    # are still populated for the older layouts.)

    # Business Themes sheet
    bt = wb["Business Themes"]
    bt_rows = list(bt.iter_rows(values_only=True))
    if _is_row_per_theme(bt_rows):
        # Newer workbooks: one theme per row --
        # Strategic Theme / Summary Intro / Key Points / Strategic Impact
        data["business_themes"] = _parse_theme_rows(bt_rows)
    else:
        theme = None
        for r in bt_rows:
            if not r:
                continue
            headline, supporting, commentary, impact = (list(r) + [None] * 4)[:4]
            if headline in ("Strategic Business Themes", "Headline"):
                continue
            if headline:
                if theme:
                    data["business_themes"].append(theme)
                theme = {
                    "headline": headline,
                    "supporting_instances": [supporting] if supporting else [],
                    "commentary": commentary,
                    "strategic_impact": impact,
                }
            elif theme and supporting:
                theme["supporting_instances"].append(supporting)
        if theme:
            data["business_themes"].append(theme)
        for t in data["business_themes"]:
            t["summary"] = t["commentary"]
            t["points"] = list(t["supporting_instances"])
            t["impact_points"] = _split_points(t["strategic_impact"])

    # Tech Themes sheet
    tt_ws = _find_sheet(wb, "Tech Themes")
    if tt_ws:
        tt_rows = list(tt_ws.iter_rows(values_only=True))
        if _is_row_per_theme(tt_rows):
            data["tech_themes"] = _parse_theme_rows(tt_rows)
        else:
            # Older layout: Headline / Commentary / Supporting Instances /
            # Strategic Impact
            header, *body = tt_rows
            for r in body:
                if not r or not r[0]:
                    continue
                headline, commentary, supporting, impact = (list(r) + [None] * 4)[:4]
                supporting_list = [
                    s.lstrip("\u2022").strip() for s in (supporting or "").split("\n") if s.strip()
                ]
                data["tech_themes"].append({
                    "headline": headline,
                    "commentary": commentary,
                    "supporting_instances": supporting_list,
                    "strategic_impact": impact,
                    "summary": supporting_list[0] if supporting_list else None,
                    "points": _split_bulleted(commentary),
                    "impact_points": _split_bulleted(impact),
                })

    # Sustainability & ESG Themes sheet (new): same row-per-theme layout as
    # Business / Tech Themes
    esg_ws = _find_sheet(wb, "Sustainability")
    if esg_ws:
        data["esg_themes"] = _parse_theme_rows(list(esg_ws.iter_rows(values_only=True)))

    # Challenges & Growth Opportunities sheet (new): Theme / Key Challenges /
    # Opportunities, bullets inside each cell
    cg_ws = _find_sheet(wb, "Challenges")
    if cg_ws:
        cg_rows = list(cg_ws.iter_rows(values_only=True))
        if cg_rows:
            hdr = (list(cg_rows[0]) + [None] * 3)[:3]
            data["challenge_headers"] = (hdr[1] or "Key Challenges", hdr[2] or "Opportunities")
        for r in cg_rows[1:]:
            if not r or not r[0]:
                continue
            theme_name, challenges, opportunities = (list(r) + [None] * 3)[:3]
            data["challenges"].append({
                "theme": theme_name,
                "challenges": _split_points(challenges),
                "opportunities": _split_points(opportunities),
            })

    # Key Triggers & Recent Developments sheet (its own sheet with clean
    # Date/Headline/Description/Source Link/Severity/Category columns, vs.
    # older workbooks where it was free text buried in Overview). Dates may
    # be real dates (older workbooks) or "October 2026"-style text (newer).
    kt_ws = _find_sheet(wb, "Key Triggers")
    if kt_ws:
        # iterate actual cells (not values_only) so we can pull the Source
        # Link hyperlink's target URL off column D, not just its display text
        for row in kt_ws.iter_rows(min_row=2, max_col=6):
            cells = list(row) + [None] * 6
            date_cell, headline_cell, desc_cell, source_cell, severity_cell, category_cell = cells[:6]
            if date_cell is None or date_cell.value is None:
                continue
            source_url = source_cell.hyperlink.target if (source_cell is not None and source_cell.hyperlink) else None
            if not source_url and source_cell is not None and isinstance(source_cell.value, str) \
                    and source_cell.value.strip().lower().startswith(("http://", "https://")):
                # newer workbooks may hold the raw URL as the cell text
                source_url = source_cell.value.strip()
            data["recent_developments"].append({
                "date": date_cell.value,
                "sort_date": _to_datetime(date_cell.value),
                "headline": headline_cell.value if headline_cell is not None else None,
                "description": desc_cell.value if desc_cell is not None else None,
                "severity": severity_cell.value if severity_cell is not None else None,
                "category": category_cell.value if category_cell is not None else None,
                "source_url": source_url,
            })
        # newest first; list.sort is stable, so events in the same month keep
        # the order they have in the workbook
        data["recent_developments"].sort(key=lambda d: d["sort_date"] or datetime.min, reverse=True)

    return data


def _is_row_per_theme(rows):
    """True for the newer theme-sheet layout (header row contains 'Key
    Points'; one theme per row, bullets inside the cells)."""
    for r in rows[:3]:
        if r and any(isinstance(c, str) and c.strip().lower() == "key points" for c in r):
            return True
    return False


def _parse_theme_rows(rows):
    """Newer theme sheets (Business / Tech / Sustainability & ESG):
    Theme | Summary Intro | Key Points | Strategic Impact -- one theme per row,
    with '\u2022 ' bulleted lines inside the Key Points / Strategic Impact cells."""
    themes = []
    for r in rows[1:]:
        if not r or not r[0]:
            continue
        headline, summary, points, impact = (list(r) + [None] * 4)[:4]
        themes.append({
            "headline": headline,
            "summary": summary,
            "points": _split_points(points),
            "impact_points": _split_points(impact),
            # legacy keys, kept so nothing that still reads them breaks
            "supporting_instances": _split_points(points),
            "commentary": summary,
            "strategic_impact": impact,
        })
    return themes


def _to_datetime(v):
    """datetime passes through; 'October 2026' / 'Oct 2026' / 'Oct-2026' ->
    datetime(2026, 10, 1); anything else -> None."""
    if isinstance(v, datetime):
        return v
    if not isinstance(v, str):
        return None
    s = v.strip().replace("-", " ").replace(",", " ")
    for fmt in ("%B %Y", "%b %Y", "%B %d %Y", "%b %d %Y", "%Y %m", "%Y"):
        try:
            return datetime.strptime(" ".join(s.split()), fmt)
        except ValueError:
            continue
    return None


# --------------------------------------------------------------------------
# 2. PPTX helpers
# --------------------------------------------------------------------------

def set_run_text(shape, text):
    """Set text into a shape's first run, clearing any other runs/paragraphs
    so formatting from that first run is preserved and nothing duplicates."""
    tf = shape.text_frame
    first_para = tf.paragraphs[0]
    if not first_para.runs:
        first_para.add_run()
    first_para.runs[0].text = text
    for extra_run in first_para.runs[1:]:
        extra_run.text = ""
    for para in tf.paragraphs[1:]:
        for run in para.runs:
            run.text = ""


def set_cell_text(cell, text, hyperlink=None):
    """Same idea as set_run_text but for a table cell. Pass `hyperlink` (a
    URL string) to turn the resulting run into a clickable link; omit/None
    leaves the cell as plain text."""
    set_run_text(cell, text)
    run = cell.text_frame.paragraphs[0].runs[0]
    run.hyperlink.address = hyperlink


def _fit_text(shape, bottom_in, min_pt=7.5):
    """Shrink a text box's font (in 0.5pt steps, never below `min_pt`) until
    its text is estimated to end above `bottom_in` (slide inches). Several
    template boxes are fixed-size with autofit off, and the newer workbooks
    carry longer bullets than the placeholders did, so without this the text
    spills over the label / panel edge underneath. Boxes whose text already
    fits are left at their template size. The height is a character-count
    estimate (Arial-ish), deliberately a touch conservative."""
    E = 914400.0
    paras = shape.text_frame.paragraphs
    texts = ["".join(r.text for r in p.runs) for p in paras]
    texts = [t for t in texts if t.strip()]
    if not texts:
        return
    base = next((r.font.size.pt for p in paras for r in p.runs if r.font.size), 10.0)
    avail = bottom_in - shape.top / E - 0.15          # top/bottom insets
    usable_w = shape.width / E - 0.2 - 0.25           # side insets + bullet indent

    def needed(pt):
        chars_per_line = max(10, usable_w * 72 / (pt * 0.5))
        lines = sum(max(1, -(-len(t) // int(chars_per_line))) for t in texts)
        return (lines * 1.2 * pt + (len(texts) - 1) * 0.5 * pt) / 72

    pt = base
    while needed(pt) > avail and pt - 0.5 >= min_pt:
        pt -= 0.5
    if pt < base:
        for p in paras:
            for r in p.runs:
                r.font.size = Pt(pt)


def set_bullet_list(shape, lines, bottom_in=None):
    """Fill a text box with one bullet paragraph per line, cloning the
    placeholder's first paragraph (bullet glyph, font, size, etc.) for each
    extra line instead of resetting to default formatting. Pass `bottom_in`
    (slide inches) to shrink the font if the text would run past that line."""
    _set_bullet_list(shape, lines)
    if bottom_in is not None:
        _fit_text(shape, bottom_in)


def _set_bullet_list(shape, lines):
    tf = shape.text_frame
    first_para = tf.paragraphs[0]
    if not first_para.runs:
        first_para.add_run()
    lines = [l for l in lines if l] or [""]
    first_para.runs[0].text = lines[0]
    for r in first_para.runs[1:]:
        r.text = ""
    for para in tf.paragraphs[1:]:
        para._p.getparent().remove(para._p)

    anchor = first_para._p
    for line in lines[1:]:
        new_p = copy.deepcopy(first_para._p)
        runs = new_p.findall(qn("a:r"))
        for extra in runs[1:]:
            new_p.remove(extra)
        runs[0].find(qn("a:t")).text = line
        anchor.addnext(new_p)
        anchor = new_p


def set_para_last_run(shape, para_idx, text):
    """Set the *last* run of a paragraph, preserving any earlier static-label
    run (e.g. the 'TCV: ' in 'TCV: ' + 'xxxxx')."""
    para = shape.text_frame.paragraphs[para_idx]
    if not para.runs:
        return
    para.runs[-1].text = text
    for run in para.runs[:-1]:
        pass  # keep static label run untouched


def set_para_full_text(shape, para_idx, label, value):
    """For paragraphs stored as one run containing 'Label: placeholder',
    rebuild as 'Label: value'."""
    para = shape.text_frame.paragraphs[para_idx]
    if not para.runs:
        return
    para.runs[0].text = f"{label}: {value}"
    for run in para.runs[1:]:
        run.text = ""


def group_path(shapes, path):
    """Walk a list of indices through nested groups/shapes."""
    s = shapes[path[0]]
    for idx in path[1:]:
        s = s.shapes[idx]
    return s


# --------------------------------------------------------------------------
# 3. Slide-by-slide fill
# --------------------------------------------------------------------------

def fmt_money(v):
    return v if v else "N/A"


def fill_slide1(slide, data, company):
    set_run_text(slide.shapes[1], company)


def _split_numbered(text):
    """'1. Foo. 2. Bar. 3. Baz.' (as stored in the Business Themes sheet's
    Strategic Impact cells) -> ['Foo.', 'Bar.', 'Baz.']"""
    if not text:
        return []
    parts = re.split(r"(?:^|\n|\s)\d+\.\s*", text.strip())
    return [p.strip() for p in parts if p.strip()]


def _split_bulleted(text):
    """'• Foo\\n• Bar' (as stored in the Tech Themes sheet's cells) ->
    ['Foo', 'Bar']"""
    if not text:
        return []
    return [ln.strip().lstrip("\u2022").strip() for ln in text.split("\n") if ln.strip()]


def _split_points(text):
    """Split a multi-point cell into a list of points, whichever style the
    workbook uses: '• Foo\\n• Bar' (newer sheets), '1. Foo. 2. Bar.' (older
    Business Themes impact cells), or a single plain sentence."""
    if not text:
        return []
    text = str(text)
    # numbered first: older Strategic Impact cells hold '1. Foo\n2. Bar' (with
    # or without line breaks) and the numbers must not survive into the deck
    if re.match(r"\s*1\.\s", text):
        return _split_numbered(text)
    if "•" in text or "\n" in text:
        return _split_bulleted(text)
    return [text.strip()]


def _split_sentences(text):
    """Split a prose paragraph into sentences for bullet display. Only
    splits on a period followed by a capitalized word, and only when the
    character before the period isn't itself a digit run like '44.8%' or
    '1,604.91', so numbers inside a sentence don't get chopped up."""
    if not text:
        return []
    parts = re.split(r"(?<=[a-zA-Z%\)])\.\s+(?=[A-Z])", text.strip())
    out = []
    for p in parts:
        p = p.strip()
        if not p:
            continue
        if not p.endswith((".", "!", "?")):
            p += "."
        out.append(p)
    return out


def fill_slide2(slide, data, company):
    shapes = slide.shapes
    desc = data["business_description"] or FLAG

    set_run_text(shapes[1], f"Business profile, key facts, and leadership overview for {company}")
    founded_match = re.search(r"(?:incorporated|founded|established) in (\d{4})", desc)
    founded = founded_match.group(1) if founded_match else FLAG

    # revenues / employees / headquarter: best-effort extraction from the
    # description (wording varies between workbooks, so a few phrasings are
    # accepted; anything not found is flagged / N/A, never guessed)
    rev_match = re.search(r"consolidated income of approximately\s*([₹$][^,\.]+(?:crore|million|billion)?)", desc)
    revenue = rev_match.group(1).strip() if rev_match else FLAG
    emp_match = re.search(
        r"employs\s+(?:approximately|roughly|about|around|over|nearly|more than)?\s*"
        r"([\d,]+ to [\d,]+|\~?[\d,]+)\s*(?:people|employees)", desc)
    employees = emp_match.group(1).strip() if emp_match else FLAG
    hq_match = re.search(r"[Hh]eadquartered in ([A-Z][A-Za-z\.\- ]+(?:, [A-Z][A-Za-z\.\- ]+)?)", desc)
    headquarter = hq_match.group(1).strip() if hq_match else "N/A"

    set_para_full_text(shapes[9], 0, "Founded (Year)", founded)
    set_para_full_text(shapes[16], 0, "Ownership", "N/A")
    set_para_full_text(shapes[18], 0, "Headquarter", headquarter)
    set_para_full_text(shapes[19], 0, "Revenues", revenue)
    set_para_full_text(shapes[20], 0, "Net Income", "N/A")
    set_para_full_text(shapes[21], 0, "Employees", employees)
    set_para_full_text(shapes[22], 0, "Market Capitalization", "N/A")

    competitor_names = [p["name"] for p in data["peers"] if p.get("name")]
    if competitor_names:
        set_bullet_list(shapes[26], competitor_names)
    else:
        set_run_text(shapes[26], FLAG)

    desc_bullets = _split_sentences(desc) if desc != FLAG else []
    if desc_bullets:
        set_bullet_list(shapes[30], desc_bullets, bottom_in=6.5)
    else:
        set_run_text(shapes[30], FLAG)


def _group_years_into_slots(years, n_slots):
    """Turn a sorted list of years into exactly n_slots groups, merging the
    closest-together years first, so every slot ends up with real data
    instead of being left empty. Returns a list of (label, [years]) tuples
    in chronological order."""
    groups = [[y] for y in sorted(years)]
    while len(groups) > n_slots:
        gaps = [groups[i + 1][0] - groups[i][-1] for i in range(len(groups) - 1)]
        merge_at = gaps.index(min(gaps))
        groups[merge_at] = groups[merge_at] + groups[merge_at + 1]
        del groups[merge_at + 1]
    labels = []
    for g in groups:
        labels.append(f"{g[0]}" if len(g) == 1 else f"{g[0]}-{g[-1]}")
    return list(zip(labels, groups))


def fill_slide3(slide, data, company):
    shapes = slide.shapes
    tl = data["timeline"]

    # year-label shapes and their matching description shapes, in
    # chronological (01->07) slide order
    year_label_shapes = [shapes[3], shapes[6], shapes[9], shapes[11], shapes[15], shapes[17], shapes[31]]
    desc_shapes = [shapes[34], shapes[35], shapes[36], shapes[37], shapes[38], shapes[39], shapes[40]]

    subtitle = f"{company}'s key milestones over time"
    set_run_text(shapes[1], subtitle)

    n_slots = len(year_label_shapes)
    years = sorted(tl.keys())
    if not years:
        for label_shape, desc_shape in zip(year_label_shapes, desc_shapes):
            set_run_text(desc_shape, FLAG)
        return

    slots = _group_years_into_slots(years, n_slots)
    pad = n_slots - len(slots)
    for _ in range(pad):
        slots.insert(0, (FLAG, []))

    def truncate(text, max_chars=110):
        if len(text) <= max_chars:
            return text
        cut = text[:max_chars].rsplit(" ", 1)[0]
        return cut + "\u2026"

    for (label, yrs), label_shape, desc_shape in zip(slots, year_label_shapes, desc_shapes):
        set_run_text(label_shape, label)
        if not yrs:
            set_run_text(desc_shape, FLAG)
        elif len(yrs) == 1:
            set_run_text(desc_shape, tl[yrs[0]])
        else:
            # multiple years sharing one slot: a short clause per year so the
            # combined text still fits the box
            per_year_budget = max(60, 260 // len(yrs))
            set_run_text(desc_shape, " ".join(truncate(tl[y], per_year_budget) for y in yrs))


def fill_slide4(slide, data, company):
    shapes = slide.shapes
    order = data["segment_order"]

    set_run_text(shapes[4], f"{company}'s core business segments and offerings")

    pairs = [(8, 9), (14, 15), (19, 20)]
    for i, (name_idx, desc_idx) in enumerate(pairs):
        if i < len(order):
            name = order[i]
            set_run_text(shapes[name_idx], name)
            sentences = _split_sentences(data["segments"][name])
            if sentences:
                set_bullet_list(shapes[desc_idx], sentences)
            else:
                set_run_text(shapes[desc_idx], FLAG)
        else:
            set_run_text(shapes[name_idx], FLAG)
            set_run_text(shapes[desc_idx], FLAG)


def fill_slide5(slide, data, company):
    shapes = slide.shapes

    set_run_text(shapes[3], f"{company}'s IT landscape: platforms, vendors, and technology focus areas")

    vendor_names = []
    for v in data["it_vendors"]:
        name = v.get("Vendor Name")
        if name and name not in vendor_names:
            vendor_names.append(name)
    if vendor_names:
        set_bullet_list(shapes[14], vendor_names)
    else:
        set_run_text(shapes[14], FLAG)

    focus_headlines = [t["headline"] for t in data["tech_themes"] if t.get("headline")]
    if focus_headlines:
        set_bullet_list(shapes[8], focus_headlines)
    else:
        set_run_text(shapes[8], FLAG)


def fill_slide6(slide, data, company):
    shapes = slide.shapes

    groups = {}
    order = []
    for v in data["it_vendors"]:
        cat = v.get("Level 2 Category") or FLAG
        if cat not in groups:
            groups[cat] = []
            order.append(cat)
        if v.get("Vendor Name") and v["Vendor Name"] not in groups[cat]:
            groups[cat].append(v["Vendor Name"])

    # (title shape path, content shape index) for the 5 category slots
    # built into this template
    slots = [
        ([6, 2, 1], 10),
        ([7, 1, 1], 11),
        ([8, 1, 1], 12),
        ([13, 2], 14),
        ([16, 2], 17),
    ]

    for i, (title_path, content_idx) in enumerate(slots):
        title_shape = group_path(shapes, title_path)
        content_shape = shapes[content_idx]
        if i < len(order):
            cat = order[i]
            set_run_text(title_shape, cat)
            names = groups[cat]
            if len(names) > 1:
                set_bullet_list(content_shape, names)
            else:
                set_run_text(content_shape, names[0] if names else FLAG)
        else:
            # fewer categories than slots on this slide: leave the extra
            # slot blank rather than showing a stale/hardcoded category name
            set_run_text(title_shape, "")
            set_run_text(content_shape, FLAG)


def _is_empty_para(para):
    return not any(r.text.strip() for r in para.runs)


def _set_first_run(para, text):
    if para.runs:
        para.runs[0].text = text
        for r in para.runs[1:]:
            r.text = ""


def _vendor_kind(vendor):
    """'software' / 'services' / None, from the Level 1 Category. Matches on
    the keyword so 'Cloud Software', 'Software', 'Cloud Services', 'Services'
    etc. all land on the right slide family."""
    l1 = str(vendor.get("Level 1 Category") or "").lower()
    if "service" in l1:
        return "services"
    if "software" in l1:
        return "software"
    return None


def _shorten(text, max_chars=400):
    """Trim a long Deal Description to whole sentences that fit a vendor card
    (the newer workbooks carry full multi-paragraph write-ups)."""
    text = " ".join(str(text or "").split())
    if len(text) <= max_chars:
        return text
    out = ""
    for s in re.split(r"(?<=[\.\!\?])\s+", text):
        if out and len(out) + 1 + len(s) > max_chars:
            break
        out = f"{out} {s}".strip()
    if len(out) > max_chars:  # a single very long sentence: cut on a word
        out = out[:max_chars].rsplit(" ", 1)[0].rstrip(",;:") + "…"
    return out


def _fmt_duration(vendor):
    """Contract duration from the 'Announced Date' -> 'End Date' columns
    (newer workbooks), e.g. '5 years' / '8 months'. 'N/A' if either is
    missing or isn't a real date."""
    start, end = vendor.get("Announced Date"), vendor.get("End Date")
    if not (isinstance(start, datetime) and isinstance(end, datetime)) or end < start:
        return "N/A"
    months = (end.year - start.year) * 12 + (end.month - start.month)
    if months >= 12:
        years = months / 12
        label = f"{years:.0f}" if months % 12 == 0 else f"{years:.1f}"
        return f"{label} year{'s' if label != '1' else ''}"
    if months >= 1:
        return f"{months} month{'s' if months != 1 else ''}"
    return "N/A"


def _fill_vendor_card(card, vendor):
    paras = card.text_frame.paragraphs
    name = vendor["Vendor Name"] if vendor else FLAG
    focus = (vendor.get("Deal Focus") or "N/A") if vendor else "N/A"
    l3 = (vendor.get("Level 3 Category") or "N/A") if vendor else FLAG
    tcv = (vendor.get("Deal Value") or "N/A") if vendor else "N/A"
    duration = _fmt_duration(vendor) if vendor else "N/A"
    desc = _shorten(vendor.get("Deal Description")) if vendor else FLAG

    joined = ["".join(r.text for r in p.runs).strip().lower() for p in paras]
    tcv_idx = next((i for i, t in enumerate(joined) if t.startswith("tcv")), None)
    duration_idx = next((i for i, t in enumerate(joined) if t.startswith("duration")), None)

    # paragraph 0: vendor name (first run), plus an optional inline
    # subtitle on a second run (some cards fold the category tag into the
    # same paragraph as the name, on its own line)
    if paras[0].runs:
        paras[0].runs[0].text = name
        if len(paras[0].runs) > 1:
            paras[0].runs[1].text = l3
            for r in paras[0].runs[2:]:
                r.text = ""

    if tcv_idx is not None:
        pre_tcv_candidates = [i for i in range(1, tcv_idx) if not _is_empty_para(paras[i])]
        if len(pre_tcv_candidates) >= 2:
            # first placeholder line = Level 3 Category subtitle,
            # last (closest to TCV) = Deal Focus tag; anything in between
            # is cleared since the template isn't consistent card to card
            subtitle_idx, tag_idx = pre_tcv_candidates[0], pre_tcv_candidates[-1]
            for i in pre_tcv_candidates[1:-1]:
                for r in paras[i].runs:
                    r.text = ""
            _set_first_run(paras[subtitle_idx], l3)
            _set_first_run(paras[tag_idx], focus)
        elif len(pre_tcv_candidates) == 1:
            _set_first_run(paras[pre_tcv_candidates[0]], focus)

        set_para_full_text(card, tcv_idx, "TCV", tcv)
    if duration_idx is not None:
        set_para_full_text(card, duration_idx, "Duration", duration)

    # description is the last non-empty paragraph in the whole box
    non_empty_after = [i for i in range(len(paras)) if not _is_empty_para(paras[i])]
    if non_empty_after:
        desc_idx = non_empty_after[-1]
        if duration_idx is None or desc_idx > duration_idx:
            _set_first_run(paras[desc_idx], desc)


def _dedup_vendors(vendors):
    by_vendor = {}
    for v in vendors:
        by_vendor.setdefault(v["Vendor Name"], v)
    return list(by_vendor.values())


def _retitle_vendor_slide(slide, title_idx, base_title, slide_num, total_slides):
    shape = slide.shapes[title_idx]
    if total_slides > 1:
        set_run_text(shape, f"{base_title} ({slide_num}/{total_slides})")
        # the extra "(x/y)" can push the title onto a second line at its
        # original size; nudge the font down a touch so it keeps fitting
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                if run.font.size:
                    run.font.size = Pt(run.font.size.pt - 3)
    else:
        set_run_text(shape, base_title)


def fill_software_vendor_slides(prs, base_slide_index, data):
    """Fill the 'IT Vendors - Software' slide's 4 cards (grouped shape) from
    every vendor whose Level 1 Category is 'Software', auto-duplicating the
    slide if there are more than 4 such vendors."""
    vendor_list = _dedup_vendors([v for v in data["it_vendors"] if _vendor_kind(v) == "software"])
    chunks = [vendor_list[i:i + 4] for i in range(0, len(vendor_list), 4)] or [[]]
    total = len(chunks)

    base_slide = prs.slides[base_slide_index]
    base_title = re.sub(r"\s*\(\d+/\d+\)\s*$", "", base_slide.shapes[3].text_frame.text).strip()

    def fill_one(slide, chunk):
        grp = slide.shapes[6]
        for i, card in enumerate(grp.shapes):
            _fill_vendor_card(card, chunk[i] if i < len(chunk) else None)

    fill_one(base_slide, chunks[0])

    for extra_i, chunk in enumerate(chunks[1:], start=2):
        insert_index = base_slide_index + extra_i - 1
        new_slide = duplicate_slide(prs, base_slide, insert_index)
        fill_one(new_slide, chunk)
        _retitle_vendor_slide(new_slide, 3, base_title, extra_i, total)

    # retitle the base slide last: copies are cloned from it, and retitling
    # shrinks the font, so doing it first would shrink the copies twice
    _retitle_vendor_slide(base_slide, 3, base_title, 1, total)

    return total


def fill_ict_vendor_slides(prs, base_slide_index, data):
    """Same idea as fill_software_vendor_slides, for the 3-card
    'IT Vendors - ICT Services' slide and vendors whose Level 1 Category
    is 'Services'."""
    vendor_list = _dedup_vendors([v for v in data["it_vendors"] if _vendor_kind(v) == "services"])
    chunks = [vendor_list[i:i + 3] for i in range(0, len(vendor_list), 3)] or [[]]
    total = len(chunks)
    card_indices = [6, 8, 9]

    base_slide = prs.slides[base_slide_index]
    base_title = re.sub(r"\s*\(\d+/\d+\)\s*$", "", base_slide.shapes[3].text_frame.text).strip()

    def fill_one(slide, chunk):
        for i, idx in enumerate(card_indices):
            _fill_vendor_card(slide.shapes[idx], chunk[i] if i < len(chunk) else None)

    fill_one(base_slide, chunks[0])

    for extra_i, chunk in enumerate(chunks[1:], start=2):
        insert_index = base_slide_index + extra_i - 1
        new_slide = duplicate_slide(prs, base_slide, insert_index)
        fill_one(new_slide, chunk)
        _retitle_vendor_slide(new_slide, 3, base_title, extra_i, total)

    _retitle_vendor_slide(base_slide, 3, base_title, 1, total)

    return total


def fill_slide9(slide, data, company):
    # kept for backwards compatibility; superseded by fill_business_theme_slides
    _fill_theme_slide(slide, data["business_themes"][:3], 1, 1, company)


_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _remap_rids(element, rid_map):
    for el in element.iter():
        for attr in list(el.attrib):
            if attr.startswith("{" + _R_NS + "}") and el.attrib[attr] in rid_map:
                el.attrib[attr] = rid_map[el.attrib[attr]]


def duplicate_slide(prs, source_slide, insert_index):
    """Duplicate a slide (its shapes and any images/media they reference)
    and move the copy to `insert_index` in the deck. This is how extra
    'Core Business Themes' slides are created on the fly when a company has
    more than 3 themes -- the template only ships with one such slide."""
    layout = source_slide.slide_layout
    new_slide = prs.slides.add_slide(layout)

    # start from a blank canvas: drop the placeholder shapes python-pptx
    # auto-added from the layout
    for shp in list(new_slide.shapes):
        shp._element.getparent().remove(shp._element)

    # copy every relevant relationship (mainly images) so copied shape XML
    # can be repointed at equivalents that exist on the new slide
    rid_map = {}
    for rel_id, rel in source_slide.part.rels.items():
        if rel.reltype in (RT.SLIDE_LAYOUT, RT.NOTES_SLIDE):
            continue
        target = rel.target_ref if rel.is_external else rel.target_part
        new_rid = new_slide.part.relate_to(target, rel.reltype, is_external=rel.is_external)
        rid_map[rel_id] = new_rid

    for shp in source_slide.shapes:
        new_el = copy.deepcopy(shp._element)
        _remap_rids(new_el, rid_map)
        new_slide.shapes._spTree.append(new_el)

    # move the newly-added slide (currently last) to the requested position
    xml_slides = prs.slides._sldIdLst
    slide_elements = list(xml_slides)
    xml_slides.remove(slide_elements[-1])
    xml_slides.insert(insert_index, slide_elements[-1])

    return new_slide


def _fill_theme_slide(slide, chunk, slide_num, total_slides, company):
    shapes = slide.shapes
    title_shape = shapes[3]
    new_title = re.sub(r"\(\d+/\d+\)", f"({slide_num}/{total_slides})", title_shape.text_frame.text)
    set_run_text(title_shape, new_title)
    set_run_text(shapes[32], f"Key business themes shaping {company}'s strategic priorities")

    columns = [
        {"headline": 33, "short": 34, "body": 35, "impact_path": [36, 0, 1]},
        {"headline": 15, "short": 16, "body": 27, "impact_path": [24, 0, 1]},
        {"headline": 22, "short": 23, "body": 28, "impact_path": [25, 0, 1]},
    ]
    for i, col in enumerate(columns):
        impact_shape = group_path(shapes, col["impact_path"])
        if i >= len(chunk):
            # fewer than 3 themes on this (last) slide: leave the extra
            # column's text empty rather than showing "xxxx"/stale content
            set_run_text(shapes[col["headline"]], "")
            set_run_text(shapes[col["short"]], "")
            set_run_text(shapes[col["body"]], "")
            set_run_text(impact_shape, "")
            continue
        t = chunk[i]
        set_run_text(shapes[col["headline"]], t["headline"] or "")
        set_run_text(shapes[col["short"]], t["summary"] or FLAG)
        if t["points"]:
            set_bullet_list(shapes[col["body"]], t["points"], bottom_in=5.5)
        else:
            set_run_text(shapes[col["body"]], FLAG)
        impact_points = t["impact_points"]
        if impact_points:
            set_bullet_list(impact_shape, impact_points, bottom_in=7.05)
        else:
            set_run_text(impact_shape, FLAG)


def fill_business_theme_slides(prs, base_slide_index, data, company):
    """Fill the template's one 'Core Business Themes' slide with the first
    3 themes, then auto-duplicate it (right after itself, before Key
    Developments) for every additional group of up to 3 themes, renumbering
    each slide's '(x/N)' title as it goes. No manual slide-adding needed."""
    themes = data["business_themes"]
    chunks = [themes[i:i + 3] for i in range(0, len(themes), 3)] or [[]]
    total = len(chunks)

    base_slide = prs.slides[base_slide_index]
    _fill_theme_slide(base_slide, chunks[0], 1, total, company)

    for extra_i, chunk in enumerate(chunks[1:], start=2):
        insert_index = base_slide_index + extra_i - 1
        new_slide = duplicate_slide(prs, base_slide, insert_index)
        _fill_theme_slide(new_slide, chunk, extra_i, total, company)

    return total


def _clean_bullets(text):
    """The Tech Themes sheet bakes '• ' markers into its multi-line cells,
    but the template boxes already render their own bullet glyph -- strip
    the workbook's markers so we don't end up with double bullets."""
    if not text:
        return text
    lines = [ln.strip().lstrip("\u2022").strip() for ln in text.split("\n") if ln.strip()]
    return " ".join(lines)


def _fill_tech_theme_slide(slide, chunk, slide_num, total_slides, company):
    shapes = slide.shapes
    title_shape = shapes[2]
    new_title = re.sub(r"\(\d+/\d+\)", f"({slide_num}/{total_slides})", title_shape.text_frame.text)
    set_run_text(title_shape, new_title)
    set_run_text(shapes[10], f"Key technology themes shaping {company}'s digital roadmap")

    columns = [
        {"headline": 8, "short": 9, "body": 11, "impact": 12},
        {"headline": 18, "short": 19, "body": 20, "impact": 21},
    ]
    for i, col in enumerate(columns):
        if i >= len(chunk):
            for key in col:
                set_run_text(shapes[col[key]], "")
            continue
        t = chunk[i]
        set_run_text(shapes[col["headline"]], t["headline"] or "")
        set_run_text(shapes[col["short"]], t["summary"] or FLAG)

        if t["points"]:
            set_bullet_list(shapes[col["body"]], t["points"], bottom_in=5.15)
        else:
            set_run_text(shapes[col["body"]], FLAG)

        impact_points = t["impact_points"]
        if impact_points:
            set_bullet_list(shapes[col["impact"]], impact_points, bottom_in=6.85)
        else:
            set_run_text(shapes[col["impact"]], FLAG)


def fill_tech_theme_slides(prs, base_slide_index, data, company):
    """Same idea as fill_business_theme_slides, but for the 'Technology
    Themes' slide, which only fits 2 themes per slide."""
    themes = data["tech_themes"]
    chunks = [themes[i:i + 2] for i in range(0, len(themes), 2)] or [[]]
    total = len(chunks)

    base_slide = prs.slides[base_slide_index]
    _fill_tech_theme_slide(base_slide, chunks[0], 1, total, company)

    for extra_i, chunk in enumerate(chunks[1:], start=2):
        insert_index = base_slide_index + extra_i - 1
        new_slide = duplicate_slide(prs, base_slide, insert_index)
        _fill_tech_theme_slide(new_slide, chunk, extra_i, total, company)

    return total


def _fill_esg_slide(slide, chunk, slide_num, total_slides, company):
    """'Sustainability Themes' slide: 2 themes per slide, each with a header,
    a bullet list (Key Points) and a Strategic Impact box. The template has
    no slot for the sheet's 'Summary Intro', so that column is not shown."""
    shapes = slide.shapes
    title_shape = shapes[3]
    new_title = re.sub(r"\(\d+/\d+\)", f"({slide_num}/{total_slides})", title_shape.text_frame.text)
    set_run_text(title_shape, new_title)
    set_run_text(shapes[6], f"Key sustainability and ESG themes shaping {company}'s responsible-business agenda")

    # the right-hand column's shapes sit inside a nested group in the template
    columns = [
        {"headline": [9], "body": [10], "impact": [13]},
        {"headline": [14, 2], "body": [14, 3], "impact": [15]},
    ]
    for i, col in enumerate(columns):
        targets = {k: group_path(shapes, path) for k, path in col.items()}
        if i >= len(chunk):
            for shp in targets.values():
                set_run_text(shp, "")
            continue
        t = chunk[i]
        set_run_text(targets["headline"], t["headline"] or "")
        if t["points"]:
            set_bullet_list(targets["body"], t["points"], bottom_in=5.2)
        else:
            set_run_text(targets["body"], FLAG)
        if t["impact_points"]:
            set_bullet_list(targets["impact"], t["impact_points"], bottom_in=6.85)
        else:
            set_run_text(targets["impact"], FLAG)


def fill_esg_theme_slides(prs, base_slide_index, data, company):
    """Same idea as fill_tech_theme_slides, for the 'Sustainability Themes'
    slide (2 themes per slide, duplicated as needed)."""
    themes = data["esg_themes"]
    chunks = [themes[i:i + 2] for i in range(0, len(themes), 2)] or [[]]
    total = len(chunks)

    base_slide = prs.slides[base_slide_index]
    _fill_esg_slide(base_slide, chunks[0], 1, total, company)

    for extra_i, chunk in enumerate(chunks[1:], start=2):
        insert_index = base_slide_index + extra_i - 1
        new_slide = duplicate_slide(prs, base_slide, insert_index)
        _fill_esg_slide(new_slide, chunk, extra_i, total, company)

    return total


def _fill_challenge_slide(slide, item, slide_num, total_slides, company, headers):
    """'Challenges & Growth Prospects' slide: one theme per slide -- the left
    panel lists its Key Challenges, the right panel its Opportunities. The
    panel headers use the sheet's own column headings; the theme name goes in
    the subtitle line."""
    shapes = slide.shapes
    base_title = re.sub(r"\s*\(\d+/\d+\)\s*$", "", shapes[2].text_frame.text).strip()
    _retitle_vendor_slide(slide, 2, base_title, slide_num, total_slides)
    set_run_text(shapes[4], f"{item['theme']}: challenges and growth opportunities for {company}")
    set_run_text(shapes[9], headers[0])
    set_run_text(shapes[10], headers[1])
    for idx, key in ((11, "challenges"), (12, "opportunities")):
        if item[key]:
            set_bullet_list(shapes[idx], item[key], bottom_in=6.5)
        else:
            set_run_text(shapes[idx], FLAG)


def fill_challenge_slides(prs, base_slide_index, data, company):
    """One slide per row of the 'Challenges & Growth Opportunities' sheet,
    duplicated from the template's single slide and numbered (x/N)."""
    items = data["challenges"]
    total = len(items)
    headers = data["challenge_headers"]

    # duplicate the still-blank slide first, then fill every copy, so the
    # (x/N) title shrink in _retitle_vendor_slide is applied exactly once per slide
    base_slide = prs.slides[base_slide_index]
    targets = [base_slide] + [
        duplicate_slide(prs, base_slide, base_slide_index + k) for k in range(1, total)
    ]
    for n, (slide, item) in enumerate(zip(targets, items), start=1):
        _fill_challenge_slide(slide, item, n, total, company, headers)

    return total


CATEGORY_TO_TYPE = {
    "Executive & Leadership Shifts": "Org Change",
    "Tech Stack & Legal Triggers": "Risk Event",
    "Financial & Corporate Structure": "Financial",
    "Corporate Expansion & Growth": "Expansion",
    "RFP & Procurement Signals": "Procurement",
}


def _dev_type(category):
    return CATEGORY_TO_TYPE.get(category, category or "Update")


def _fmt_date(d):
    if isinstance(d, datetime):
        return f"{d.strftime('%b')} {d.day}, {d.year}"
    parsed = _to_datetime(d)
    if parsed:  # month-level text like "October 2026" -> "Oct 2026"
        return f"{parsed.strftime('%b')} {parsed.year}"
    return str(d) if d else "N/A"


def _fill_key_dev_table(slide, chunk):
    table = slide.shapes[6].table
    for i in range(4):
        row = i + 1
        if i < len(chunk):
            ev = chunk[i]
            set_cell_text(table.cell(row, 0), _fmt_date(ev["date"]), hyperlink=ev.get("source_url"))
            set_cell_text(table.cell(row, 1), ev["headline"] or "")
            set_cell_text(table.cell(row, 2), _dev_type(ev["category"]))
            set_cell_text(table.cell(row, 3), ev["description"] or "")
        else:
            for c in range(4):
                set_cell_text(table.cell(row, c), "")


def _fill_key_dev_slide(slide, chunk, slide_num, total_slides, all_events):
    shapes = slide.shapes
    title_shape = shapes[4]
    new_title = re.sub(r"\(\d+/\d+\)", f"({slide_num}/{total_slides})", title_shape.text_frame.text)
    set_run_text(title_shape, new_title)
    if all_events:
        summary = "; ".join(ev["headline"] for ev in all_events[:2] if ev["headline"])
    else:
        summary = FLAG
    set_run_text(shapes[5], summary)
    _fill_key_dev_table(slide, chunk)


def fill_key_development_slides(prs, base_indices, data):
    """Fill the template's existing 'Key Developments' slides (normally 2,
    4 rows each) with real events, newest first. If there are more events
    than the template has room for, extra slides are duplicated the same
    way theme slides are; if there are fewer, leftover rows are cleared
    rather than left showing stale placeholder dates/headlines."""
    events = data["recent_developments"]
    chunks = [events[i:i + 4] for i in range(0, len(events), 4)] or [[]]
    total = max(len(chunks), len(base_indices))

    for i, idx in enumerate(base_indices):
        slide = prs.slides[idx]
        chunk = chunks[i] if i < len(chunks) else []
        _fill_key_dev_slide(slide, chunk, i + 1, total, events)

    last_idx = base_indices[-1]
    for extra_i in range(len(base_indices), len(chunks)):
        insert_index = last_idx + (extra_i - len(base_indices)) + 1
        new_slide = duplicate_slide(prs, prs.slides[last_idx], insert_index)
        _fill_key_dev_slide(new_slide, chunks[extra_i], extra_i + 1, total, events)

    return total


# --------------------------------------------------------------------------
# 4. Slide lookup (by title) and template clean-up
# --------------------------------------------------------------------------

# Slide families after the fixed first six, identified by the title text on
# the slide. Order matters: the ICT slide also carries a hidden 'IT Vendors -
# Software' tab label, so ICT must be tested before Software.
_FAMILY_PATTERNS = [
    ("challenges", re.compile(r"^Challenges\s*&\s*Growth", re.I)),
    ("esg", re.compile(r"^Sustainability Themes", re.I)),
    ("tech", re.compile(r"^Technology Themes", re.I)),
    ("business", re.compile(r"^Core Business Themes", re.I)),
    ("keydev", re.compile(r"^Key Developments", re.I)),
    ("ict", re.compile(r"^IT Vendors\s*[–—-]\s*ICT Services", re.I)),
    ("software", re.compile(r"^IT Vendors\s*[–—-]\s*Software", re.I)),
]

# If the template holds several copies of the same (non-Key-Developments)
# slide, which one is the real one? The template author normally edits the
# copy they are working on and forgets the older one, so keep the last.
KEEP_WHICH_DUPLICATE = "last"


def _slide_family(slide):
    texts = [sh.text_frame.text.strip() for sh in slide.shapes if sh.has_text_frame]
    for family, pattern in _FAMILY_PATTERNS:
        if any(pattern.search(t) for t in texts):
            return family
    return None


def _family_indices(prs, family):
    """Positions (in the deck's current order) of every slide of a family."""
    return [i for i, s in enumerate(prs.slides) if _slide_family(s) == family]


def _remove_slide(prs, slide):
    """Delete a slide from the deck and renumber the remaining slide parts
    (python-pptx names new slides after the slide count, so skipping the
    renumber could hand a new slide an existing slide's file name)."""
    sld_id_lst = prs.slides._sldIdLst
    sld_id = list(sld_id_lst)[list(prs.slides).index(slide)]
    prs.part.drop_rel(sld_id.rId)
    sld_id_lst.remove(sld_id)
    prs.part.rename_slide_parts([s.rId for s in sld_id_lst])


def _drop_spare_template_slides(prs):
    """The template sometimes ships with two copies of one slide (e.g. a
    leftover 'Technology Themes' slide from editing). The generator already
    creates as many slides per family as the data needs, so keep one copy per
    family and drop the rest -- otherwise the extra one would stay behind
    with its 'xxxx' placeholders. 'Key Developments' legitimately ships as
    two base slides, so it is left alone."""
    for family, _ in _FAMILY_PATTERNS:
        if family == "keydev":
            continue
        idx = _family_indices(prs, family)
        if len(idx) < 2:
            continue
        keep = idx[-1] if KEEP_WHICH_DUPLICATE == "last" else idx[0]
        for i in sorted((j for j in idx if j != keep), reverse=True):
            _remove_slide(prs, prs.slides[i])


# --------------------------------------------------------------------------
# 5. Public entry point
# --------------------------------------------------------------------------

def generate_ppt(excel_path, pptx_path, company, output_path):
    """Fill `pptx_path` (the RefractOne template) with data parsed from
    `excel_path` for `company`, and save the result to `output_path`.

    This is the function app.py (the Streamlit UI) imports and calls
    directly, so it takes plain file paths in/out and has no CLI/argparse
    dependency. Returns a dict of slide-count stats for logging/reporting.
    """
    data = parse_excel(excel_path)
    prs = Presentation(pptx_path)
    slides = prs.slides

    # the first six slides (cover ... IT Vendor Summary) are fixed positions
    fill_slide1(slides[0], data, company)
    fill_slide2(slides[1], data, company)
    fill_slide3(slides[2], data, company)
    fill_slide4(slides[3], data, company)
    fill_slide5(slides[4], data, company)
    fill_slide6(slides[5], data, company)

    # everything after that is located by slide title, so the template can
    # gain / lose / reorder slides without this function needing to change
    _drop_spare_template_slides(prs)

    def start_of(family):
        idx = _family_indices(prs, family)
        return idx[0] if idx else None

    def fill_family(family, fill, has_data=True):
        """Run `fill(base_index)` for a template family; returns its slide
        count. A family the template doesn't have is skipped; one whose data
        the workbook doesn't have is removed rather than left as a blank
        'xxxx' slide."""
        start = start_of(family)
        if start is None:
            return 0
        if not has_data:
            _remove_slide(prs, prs.slides[start])
            return 0
        return fill(start)

    n_software = fill_family("software", lambda i: fill_software_vendor_slides(prs, i, data))
    n_ict = fill_family("ict", lambda i: fill_ict_vendor_slides(prs, i, data))
    n_business = fill_family("business", lambda i: fill_business_theme_slides(prs, i, data, company))
    n_tech = fill_family("tech", lambda i: fill_tech_theme_slides(prs, i, data, company))
    n_challenges = fill_family(
        "challenges", lambda i: fill_challenge_slides(prs, i, data, company), bool(data["challenges"]))
    n_esg = fill_family(
        "esg", lambda i: fill_esg_theme_slides(prs, i, data, company), bool(data["esg_themes"]))

    keydev_base = _family_indices(prs, "keydev")
    n_keydev = fill_key_development_slides(prs, keydev_base, data) if keydev_base else 0

    unplaced = [v["Vendor Name"] for v in data["it_vendors"] if _vendor_kind(v) is None]

    prs.save(output_path)

    return {
        "output_path": output_path,
        "n_software": n_software,
        "n_ict": n_ict,
        "n_business": n_business,
        "n_business_themes": len(data["business_themes"]),
        "n_tech": n_tech,
        "n_tech_themes": len(data["tech_themes"]),
        "n_challenges": n_challenges,
        "n_esg": n_esg,
        "n_esg_themes": len(data["esg_themes"]),
        "n_keydev": n_keydev,
        "n_events": len(data["recent_developments"]),
        "unplaced_vendors": unplaced,
    }


# --------------------------------------------------------------------------
# 5. CLI
# --------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--excel", required=True)
    ap.add_argument("--pptx", required=True)
    ap.add_argument("--output", required=True)
    ap.add_argument("--company", required=True, help="Company display name for slide 1 / timeline subtitle")
    args = ap.parse_args()

    stats = generate_ppt(args.excel, args.pptx, args.company, args.output)

    print(f"Saved: {stats['output_path']}")
    print(f"IT Vendors - Software: {stats['n_software']} slide(s).")
    print(f"IT Vendors - ICT Services: {stats['n_ict']} slide(s).")
    print(f"Business Themes: {stats['n_business']} slide(s) for {stats['n_business_themes']} theme(s).")
    print(f"Technology Themes: {stats['n_tech']} slide(s) for {stats['n_tech_themes']} theme(s).")
    print(f"Challenges & Growth Prospects: {stats['n_challenges']} slide(s).")
    print(f"Sustainability Themes: {stats['n_esg']} slide(s) for {stats['n_esg_themes']} theme(s).")
    print(f"Key Developments: {stats['n_keydev']} slide(s) for {stats['n_events']} event(s).")


if __name__ == "__main__":
    main()
