import streamlit as st
import re
import os
import json
import tempfile
import mammoth
import docx
import base64
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from bs4 import BeautifulSoup, NavigableString
from collections import Counter

# Import components for HTML embedding.
import streamlit.components.v1 as components

REFRESH_EVERY = 5

# ---- handle pending refresh ----------------------------------------
if st.session_state.get("pending_refresh", False):
    # Keys we must keep so Step 2 resumes instantly
    keep_keys = {
        "book_name",            # your title
        "docx_path",            # temp file we rebuilt earlier
        "unknown_speakers",     # remaining lines
        "step",                 # current UI step
    }

    preserved = {k: st.session_state[k] for k in keep_keys if k in st.session_state}
    st.session_state.clear()            # drop the bulky / growing stuff
    st.session_state.update(preserved)
    st.session_state.pending_refresh = False
    st.rerun()             # safe here – we're *not* in a callback
# --------------------------------------------------------------------

# Inject custom CSS
custom_css = """
<style>
:root {
  --primary-color: #008080;      /* Teal */
  --primary-hover: #007070;
  --background-color: #fdfdfd;
  --text-color: #222222;
  --card-background: #ffffff;
  --card-shadow: 0 6px 12px rgba(0, 0, 0, 0.15);
  --border-radius: 10px;
  --font-family: 'Avenir', sans-serif;
  --accent-color: #ff9900;
}

/* Global Styles */
body {
  background-color: var(--background-color);
  font-family: var(--font-family);
  color: var(--text-color);
  margin: 0;
  padding: 0;
}

h1, h2, h3, h4, h5, h6 {
  color: var(--text-color);
  font-weight: 700;
  margin-bottom: 0.5em;
}

/* Button Styles */
div.stButton > button {
  background-color: var(--primary-color);
  color: #ffffff;
  border: none;
  padding: 0.75em 1.25em;
  border-radius: var(--border-radius);
  cursor: pointer;
  transition: background-color 0.3s ease, transform 0.2s;
}

div.stButton > button:hover {
  background-color: var(--primary-hover);
  transform: translateY(-2px);
}

/* Card/Container Styling */
.custom-container {
  background: var(--card-background);
  padding: 2em;
  border-radius: var(--border-radius);
  box-shadow: var(--card-shadow);
  margin-bottom: 2em;
}

.css-1d391kg {
  background: var(--card-background);
  padding: 1em;
  border-radius: var(--border-radius);
  box-shadow: var(--card-shadow);
}

/* Form Element Styling */
input, select, textarea {
  border: 1px solid #ddd;
  border-radius: 4px;
  padding: 0.5em;
  font-size: 1em;
}

input:focus, select:focus, textarea:focus {
  outline: none;
  border-color: var(--primary-color);
  box-shadow: 0 0 5px rgba(0, 128, 128, 0.3);
}
</style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

# ---------------------------
# Global Constants & Helper Functions
# ---------------------------

COLOR_PALETTE = {
    "dark grey": (180, 178, 179, 1, "rgb(30, 28, 29)"),
    "burgundy": (207, 156, 153, 1, "rgb(55, 4, 1)"),
    "red": (255, 153, 153, 1, "rgb(107, 9, 6)"),
    "orange": (255, 198, 153, 1, "rgb(119, 62, 17)"),
    "yellow": (255, 241, 153, 1, "rgb(107, 93, 5)"),
    "dark yellow": (230, 223, 153, 1, "rgb(78, 71, 1)"),
    "brown": (205, 192, 173, 1, "rgb(61, 48, 29)"),
    "silver": (243, 243, 243, 1, "rgb(96, 96, 96)"),
    "light green": (160, 255, 153, 1, "rgb(20, 115, 13)"),
    "dark green": (157, 192, 152, 1, "rgb(7, 42, 2)"),
    "turquoise": (153, 255, 234, 1, "rgb(3, 105, 84)"),
    "light blue": (153, 248, 254, 1, "rgb(9, 104, 110)"),
    "bright blue": (152, 166, 255, 1, "rgb(4, 18, 107)"),
    "dark blue": (152, 166, 255, 1, "rgb(4, 18, 107)"),
    "navy blue": (153, 162, 201, 1, "rgb(4, 13, 52)"),
    "dark purple": (219, 173, 222, 1, "rgb(68, 22, 71)"),
    "light purple": (231, 203, 254, 1, "rgb(81, 53, 104)"),
    "bright pink": (255, 153, 236, 1, "rgb(112, 10, 93)"),
    "light pink": (254, 238, 248, 1, "rgb(103, 87, 97)"),
    "pale pink": (254, 238, 248, 1, "rgb(103, 87, 97)"),
    "wine": (234, 184, 185, 1, "rgb(90, 40, 41)"),
    "lime": (230, 244, 182, 1, "rgb(81, 95, 33)"),
    "none": (134, 8, 0, 1.0, "rgb(134, 8, 0)"),
    "error": (0, 0, 0, 0, "")  # For "Error": transparent background, no text color override.
}
SAVED_COLORS_FILE = "speaker_colors.json"
PROGRESS_FILE = "progress.json"

def normalize_text(text):
    text = text.replace("\u00A0", " ")
    text = text.replace("…", "...")
    text = text.replace("“", "\"").replace("”", "\"")
    text = text.replace("’", "'").replace("‘", "'")
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

def match_normalize(text):
    return text.replace("’", "'").replace("‘", "'")

def normalize_speaker_name(name):
    # Replace typographic apostrophes with straight ones, remove periods, lowercase, and trim.
    return name.replace("’", "'").replace("‘", "'").replace(".", "").lower().strip()

def smart_title(name):
    words = name.split()
    if not words:
        return name
    exceptions = {"ps", "pc", "ds", "di", "dci"}
    new_words = []
    for w in words:
        if w.lower() in exceptions:
            new_words.append(w.upper())
        else:
            new_words.append(w.capitalize())
    result = " ".join(new_words)
    result = re.sub(r"\(([mf])\)$", lambda m: "(" + m.group(1).upper() + ")", result, flags=re.IGNORECASE)
    return result

def write_file_atomic(filepath, lines):
    with open(filepath, "w", encoding="utf-8") as f:
        f.writelines(lines)
        f.flush()
        os.fsync(f.fileno())

# ---------------------------
# Auto-Save & Auto-Load Functions
# ---------------------------
def auto_save():
    data = {
        "step": st.session_state.get("step", 1),
        "quotes_lines": st.session_state.get("quotes_lines"),
        "speaker_colors": st.session_state.get("speaker_colors"),
        "unknown_index": st.session_state.get("unknown_index", 0),
        "console_log": st.session_state.get("console_log", []),
        "canonical_map": st.session_state.get("canonical_map"),
        "book_name": st.session_state.get("book_name"),
        "existing_speaker_colors": st.session_state.get("existing_speaker_colors")
    }
    if "docx_bytes" in st.session_state and st.session_state.docx_bytes is not None:
        data["docx_bytes"] = base64.b64encode(st.session_state.docx_bytes).decode("utf-8")
    with open(PROGRESS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4)
    if st.session_state.get("speaker_colors") is not None:
        with open(SAVED_COLORS_FILE, "w", encoding="utf-8") as f:
            json.dump(st.session_state.speaker_colors, f, indent=4, ensure_ascii=False)
    if st.session_state.get("quotes_lines") and st.session_state.get("book_name"):
        quotes_filename = f"{st.session_state.book_name}-quotes.txt"
        with open(quotes_filename, "w", encoding="utf-8") as f:
            quotes_text = "".join(st.session_state.quotes_lines)
            f.write(quotes_text)

def auto_load():
    if os.path.exists(PROGRESS_FILE):
        with open(PROGRESS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        for key, value in data.items():
            st.session_state[key] = value
        if "existing_speaker_colors" in st.session_state and st.session_state.existing_speaker_colors:
            st.session_state.existing_speaker_colors = {normalize_speaker_name(k): v for k, v in st.session_state.existing_speaker_colors.items()}
        if "docx_bytes" in st.session_state:
            docx_bytes = base64.b64decode(st.session_state["docx_bytes"].encode("utf-8"))
            st.session_state.docx_bytes = docx_bytes
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_docx:
                tmp_docx.write(docx_bytes)
                st.session_state.docx_path = tmp_docx.name

if os.path.exists(PROGRESS_FILE):
    if st.button("Load Saved Progress"):
         auto_load()
         st.rerun()

# ---------------------------
# Alternative Extraction Functions
# ---------------------------
ATTACH_NO_SPACE = {"'", "’", "‘", '"', "“", "”", ",", ".", ";", ":", "?", "!"}
DASHES = {"-", "–", "—"}

def smart_join(run_texts):
    if not run_texts:
        return ""
    result = run_texts[0]
    for text in run_texts[1:]:
        if not text:
            continue
        if result and result[-1].isspace():
            result += text.lstrip()
        elif text[0].isspace():
            result += text
        elif text.startswith("...") or text.startswith("…"):
            result = result.rstrip()
            result += text
        elif text[0] in ATTACH_NO_SPACE:
            result += text
        elif result[-1] in DASHES:
            if len(result) == 1 or result[-2].isspace():
                result += text
            else:
                if result[-2].isalnum() and text[0].isalnum():
                    result += text
                else:
                    result += " " + text
        else:
            if result[-1].isalnum() and text[0].isalnum():
                result += text
            else:
                result += " " + text
    return result

def extract_italicized_text(paragraph):
    italic_blocks = []
    current_block = []
    for run in paragraph.runs:
        if run.italic:
            current_block.append(run.text)
        else:
            joined = smart_join(current_block)
            if len(joined.split()) >= 2:
                italic_blocks.append(joined)
            current_block = []
    joined = smart_join(current_block)
    if len(joined.split()) >= 2:
        italic_blocks.append(joined)
    return italic_blocks

def extract_dialogue_from_docx(book_name, docx_path):
    doc = docx.Document(docx_path)
    quote_pattern = re.compile(r'(?:^|\s)(["“].+?["”])(?=$|\s)')
    dialogue_list = []
    line_number = 1
    for para in doc.paragraphs:
        text = para.text.strip()
        quotes = quote_pattern.findall(text)
        if quotes:
            for quote in quotes:
                dialogue_list.append(f"{line_number}. Unknown: {quote}")
                line_number += 1
        else:
            italic_texts = extract_italicized_text(para)
            for italic_text in italic_texts:
                dialogue_list.append(f"{line_number}. Unknown: {italic_text}")
                line_number += 1
    output_path = f"{book_name}-quotes.txt"
    with open(output_path, "w", encoding="utf-8") as f:
        f.write("\n".join(dialogue_list))
    return dialogue_list

# ---------------------------
# DOCX-to-HTML & Marking Functions
# ---------------------------
def prepend_marker_to_paragraph(paragraph, marker_text):
    p = paragraph._p
    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = marker_text + " "
    r.append(t)
    p.insert(0, r)

def create_marker_docx(original_docx, marker_docx):
    doc = docx.Document(original_docx)
    for idx, para in enumerate(doc.paragraphs):
        marker = f"[[[P{idx}]]]"
        prepend_marker_to_paragraph(para, marker)
    doc.save(marker_docx)

def convert_docx_to_html_mammoth(docx_file):
    with open(docx_file, "rb") as f:
        result = mammoth.convert_to_html(f)
        return result.value

def get_manual_indentation(docx_file):
    doc = docx.Document(docx_file)
    indented_paras = {}
    for idx, para in enumerate(doc.paragraphs):
        left = para.paragraph_format.left_indent
        right = para.paragraph_format.right_indent
        if (left is not None and left.pt > 0) or (right is not None and right.pt > 0):
            indented_paras[idx] = (left, right)
    return indented_paras

def convert_length_to_px(length):
    return length.pt * 1.33 if length is not None else 0

# ---------------------------
# Dialogue Highlighting Functions
# ---------------------------
def highlight_across_nodes(parent, quote, highlight_style, soup):
    full_text = parent.get_text()
    full_text_lower = match_normalize(full_text).lower()
    quote_lower = match_normalize(quote).lower()
    idx = full_text_lower.find(quote_lower)
    if idx == -1:
        return False
    end_idx = idx + len(quote)
    running_index = 0
    for descendant in list(parent.descendants):
        if isinstance(descendant, NavigableString):
            text = str(descendant)
            text_length = len(text)
            node_start = running_index
            node_end = running_index + text_length
            if node_end > idx and node_start < end_idx:
                overlap_start = max(idx, node_start)
                overlap_end = min(end_idx, node_end)
                rel_start = overlap_start - node_start
                rel_end = overlap_end - node_start
                before = text[:rel_start]
                match_text = text[rel_start:rel_end]
                after = text[rel_end:]
                new_nodes = []
                if before:
                    new_nodes.append(NavigableString(before))
                span_tag = soup.new_tag("span", attrs={"class": "highlight", "style": highlight_style})
                span_tag.string = match_text
                new_nodes.append(span_tag)
                if after:
                    new_nodes.append(NavigableString(after))
                descendant.replace_with(*new_nodes)
            running_index += text_length
    return True

def highlight_quote_in_parent(parent, quote, highlight_style, soup):
    stripped_quote = quote.strip('“”"')
    stripped_quote_lower = match_normalize(stripped_quote).lower()
    for child in parent.contents:
        if isinstance(child, NavigableString):
            child_text = str(child)
            child_text_lower = match_normalize(child_text).lower()
            idx = child_text_lower.find(stripped_quote_lower)
            if idx != -1:
                before = child_text[:idx]
                match_text = child_text[idx: idx + len(stripped_quote)]
                after = child_text[idx + len(stripped_quote):]
                new_nodes = []
                if before:
                    new_nodes.append(NavigableString(before))
                span_tag = soup.new_tag("span", attrs={"class": "highlight", "style": highlight_style})
                span_tag.string = match_text
                new_nodes.append(span_tag)
                if after:
                    new_nodes.append(NavigableString(after))
                child.replace_with(*new_nodes)
                return True
        elif hasattr(child, 'contents'):
            if highlight_quote_in_parent(child, quote, highlight_style, soup):
                return True
    return highlight_across_nodes(parent, stripped_quote, highlight_style, soup)

def build_candidate_info(soup):
    candidates = soup.find_all(['p', 'li', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6'])
    candidate_info = []
    global_offset = 0
    for candidate in candidates:
        text = candidate.get_text()
        length = len(text)
        candidate_info.append((candidate, global_offset, global_offset + length, text))
        global_offset += length
    return candidate_info

def highlight_in_candidate(candidate, quote, highlight_style, soup, start_offset=0):
    full_text = candidate.get_text()
    full_text_lower = match_normalize(full_text).lower()
    quote_lower = match_normalize(quote).lower()
    pos = full_text_lower.find(quote_lower, start_offset)
    if pos == -1:
        return None
    match_end = pos + len(quote)
    running_index = 0
    for descendant in list(candidate.descendants):
        if isinstance(descendant, NavigableString):
            text = str(descendant)
            text_length = len(text)
            node_start = running_index
            node_end = running_index + text_length
            if node_end > pos and node_start < match_end:
                overlap_start = max(pos, node_start)
                overlap_end = min(match_end, node_end)
                rel_start = overlap_start - node_start
                rel_end = overlap_end - node_start
                before = text[:rel_start]
                match_text = text[rel_start:rel_end]
                after = text[rel_end:]
                new_nodes = []
                if before:
                    new_nodes.append(NavigableString(before))
                span_tag = soup.new_tag("span", attrs={"class": "highlight", "style": highlight_style})
                span_tag.string = match_text
                new_nodes.append(span_tag)
                if after:
                    new_nodes.append(NavigableString(after))
                descendant.replace_with(*new_nodes)
            running_index += text_length
    return match_end

def highlight_dialogue_in_html(html, quotes_list, speaker_colors):
    soup = BeautifulSoup(html, "html.parser")
    candidate_info = build_candidate_info(soup)
    unmatched_quotes = []
    last_global_offset = 0
    for quote_data in quotes_list:
        expected_quote = quote_data['quote'].strip('“”"')
        expected_quote_lower = match_normalize(expected_quote).lower()
        speaker = quote_data['speaker']
        norm_speaker = normalize_speaker_name(speaker)
        color_choice = st.session_state.speaker_colors.get(norm_speaker, "none")
        if norm_speaker == "unknown":
            color_choice = "none"
        rgba = COLOR_PALETTE.get(color_choice, COLOR_PALETTE["none"])
        if color_choice == "none":
            highlight_style = f"color: rgb({rgba[0]}, {rgba[1]}, {rgba[2]}); background-color: transparent;"
        else:
            highlight_style = f"color: {rgba[4]}; background-color: rgba({rgba[0]}, {rgba[1]}, {rgba[2]}, {rgba[3]});"
        matched = False
        for candidate, start, end, text in candidate_info:
            if end < last_global_offset:
                continue
            local_start = last_global_offset - start if last_global_offset > start else 0
            candidate_text_norm = match_normalize(text).lower()
            pos = candidate_text_norm.find(expected_quote_lower, local_start)
            if pos != -1:
                match_end_local = highlight_in_candidate(candidate, quote_data['quote'], highlight_style, soup, local_start)
                if match_end_local is not None:
                    last_global_offset = start + match_end_local
                    matched = True
                    break
        if not matched:
            for candidate, start, end, text in candidate_info:
                candidate_text_norm = match_normalize(text).lower()
                pos = candidate_text_norm.find(expected_quote_lower)
                if pos != -1:
                    match_end_local = highlight_in_candidate(candidate, quote_data['quote'], highlight_style, soup, 0)
                    if match_end_local is not None:
                        if start + match_end_local > last_global_offset:
                            last_global_offset = start + match_end_local
                        matched = True
                        break
        if not matched:
            unmatched_quotes.append(f"{quote_data['speaker']}: \"{quote_data['quote']}\" [Index: {quote_data['index']}]")
    if unmatched_quotes:
        with open("unmatched_quotes.txt", "w", encoding="utf-8") as f:
            f.write("\n".join(unmatched_quotes))
        st.write(f"⚠️ Unmatched quotes saved to 'unmatched_quotes.txt' ({len(unmatched_quotes)} entries)")
    return str(soup)

def apply_manual_indentation_with_markers(original_docx, html):
    indented_paras = get_manual_indentation(original_docx)
    soup = BeautifulSoup(html, "html.parser")
    marker_regex = re.compile(r"\[\[\[P(\d+)\]\]\]")
    candidate_tags = ['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'li']
    for tag in soup.find_all(candidate_tags):
        text = tag.get_text()
        match = marker_regex.search(text)
        if match:
            para_index = int(match.group(1))
            for text_node in tag.find_all(string=True):
                if marker_regex.search(text_node):
                    new_text = marker_regex.sub("", text_node)
                    text_node.replace_with(new_text)
            if para_index in indented_paras:
                left, right = indented_paras[para_index]
                left_px = convert_length_to_px(left)
                right_px = convert_length_to_px(right)
                style_str = f"margin-left: {left_px}px; margin-right: {right_px}px;"
                if tag.has_attr("style"):
                    tag["style"] += " " + style_str
                else:
                    tag["style"] = style_str
    return str(soup)

# ---------------------------
# Summary & Ranking Functions
# ---------------------------
def generate_summary_html(quotes_list, speakers, speaker_colors):
    counts = Counter(quote["speaker"] for quote in quotes_list)
    total_lines = sum(counts.values())
    summary_order = []
    if "Unknown" in counts:
        summary_order.append("Unknown")
    for sp in speakers:
        if sp != "Unknown" and sp not in summary_order:
            summary_order.append(sp)
    for sp in counts:
        if sp not in summary_order:
            summary_order.append(sp)
    lines = []
    lines.append('<div id="character-summary" style="border: 1px solid #ccc; padding: 10px; margin-bottom: 20px;">')
    lines.append('<h2 style="margin: 0 0 5px 0;">Character Summary</h2>')
    for sp in summary_order:
        count = counts.get(sp, 0)
        percentage = round((count / total_lines) * 100) if total_lines > 0 else 0
        color_key = speaker_colors.get(normalize_speaker_name(sp), "none")
        if sp.lower() == "unknown":
            color_key = "none"
        rgba = COLOR_PALETTE.get(color_key, COLOR_PALETTE["none"])
        if color_key == "none":
            style = f"color: rgb({rgba[0]}, {rgba[1]}, {rgba[2]}); background-color: transparent;"
        else:
            style = f"color: {rgba[4]}; background-color: rgba({rgba[0]}, {rgba[1]}, {rgba[2]}, {rgba[3]});"
        lines.append(f'<p style="margin: 0; line-height: 1.2; padding: 8px 0;"><span class="highlight" style="{style}">{sp}</span> - {count} lines - {percentage}%</p>')
    lines.append('</div>')
    return "\n".join(lines)

def generate_ranking_html(quotes_list, speaker_colors):
    counts = Counter(quote["speaker"] for quote in quotes_list)
    total_lines = sum(counts.values())
    filtered = [(sp, count) for sp, count in counts.items() if sp.lower() != "unknown" and count > 1]
    filtered.sort(key=lambda x: x[1], reverse=True)
    lines = []
    lines.append('<div id="speaker-ranking" style="margin-top: 20px;">')
    lines.append('<h2 style="margin: 0 0 5px 0;">Speaker Ranking</h2>')
    for sp, count in filtered:
        percentage = round((count / total_lines) * 100) if total_lines > 0 else 0
        color_key = speaker_colors.get(normalize_speaker_name(sp), "none")
        rgba = COLOR_PALETTE.get(color_key, COLOR_PALETTE["none"])
        if color_key == "none":
            style = f"color: rgb({rgba[0]}, {rgba[1]}, {rgba[2]}); background-color: transparent;"
        else:
            style = f"color: {rgba[4]}; background-color: rgba({rgba[0]}, {rgba[1]}, {rgba[2]}, {rgba[3]});"
        lines.append(f'<p style="margin: 0; line-height: 1.2; padding: 8px 0;"><span class="highlight" style="{style}">{sp}</span> - {count} lines - {percentage}%</p>')
    lines.append('</div>')
    return "\n".join(lines)

# ---------------------------
# Canonical Speaker & Quote Functions
# ---------------------------
def get_canonical_speakers(quotes_file):
    speakers = []
    pattern = re.compile(r"^\s*\d+(?:[a-zA-Z]+)?\.\s+([^:]+):")
    with open(quotes_file, "r", encoding="utf-8") as f:
        for line in f:
            match = pattern.match(line.strip())
            if match:
                speaker_raw = match.group(1).strip()
                speakers.append(smart_title(str(speaker_raw)))  # Ensure it's a string
    seen = set()
    canonical_speakers = []
    for s in speakers:
        norm = normalize_speaker_name(str(s))
        if norm not in seen:
            seen.add(norm)
            canonical_speakers.append(s)
    canonical_map = {normalize_speaker_name(str(s)): s for s in canonical_speakers}
    return canonical_speakers, canonical_map

def load_quotes(quotes_file, canonical_map):
    quotes_list = []
    pattern = re.compile(r"^\s*([0-9]+(?:[a-zA-Z]+)?)\.\s+([^:]+):\s*(?:[“\"])?(.+?)(?:[”\"])?\s*$")
    with open(quotes_file, "r", encoding="utf-8") as f:
        for line in f:
            match = pattern.match(line.strip())
            if match:
                index, speaker_raw, quote = match.groups()
                effective = smart_title(speaker_raw)
                norm = normalize_speaker_name(effective)
                canonical = canonical_map.get(norm, effective)
                quotes_list.append({
                    "index": index,
                    "speaker": canonical,
                    "quote": quote.strip()
                })
    return quotes_list

def load_existing_colors():
  if os.path.exists(SAVED_COLORS_FILE):
    with open(SAVED_COLORS_FILE, "r", encoding="utf-8") as f:
        loaded_colors = json.load(f)
    normalized_loaded = {normalize_speaker_name(k): v for k, v in loaded_colors.items()}
    st.session_state.speaker_colors = normalized_loaded
    st.session_state.existing_speaker_colors = normalized_loaded

def save_speaker_colors(speaker_colors):
    with open(SAVED_COLORS_FILE, "w", encoding="utf-8") as f:
        json.dump(speaker_colors, f, indent=4, ensure_ascii=False)

# ---------------------------
# Restart Helper Function
# ---------------------------
def restart_app():
    st.session_state.clear()
    st.rerun()

# ---------------------------
# Streamlit Multi-Step UI
# ---------------------------
if 'step' not in st.session_state:
    st.session_state.step = 1

# ========= STEP 1: Upload & Initialize =========
if st.session_state.step == 1:
    st.markdown("<h4>DOCX to HTML Converter with Dialogue Highlighting</h4>", unsafe_allow_html=True)
    st.write("Upload your DOCX and quotes text files. Optionally, upload an existing speaker_colors.json file.")
    st.write("Alternatively, upload **just a DOCX** to create a quotes text file.")
    
    if "docx_bytes" in st.session_state:
        st.success("DOCX already uploaded and processed.")
        if st.session_state.get("docx_only", False):
            if st.session_state.get("quotes_lines") is not None:
                quotes_txt = "\n".join(st.session_state.quotes_lines)
                st.download_button("Download Extracted Quotes TXT", quotes_txt.encode("utf-8"),
                                   file_name=f"{st.session_state.book_name}-quotes.txt", mime="text/plain")
                if st.button("Restart", key="restart_docx"):
                    restart_app()
                if st.button("Continue", key="continue_docx"):
                    st.session_state.docx_only = False
                    st.session_state.unknown_index = 0
                    st.session_state.console_log = []
                    st.session_state.step = 2
                    auto_save()
                    st.rerun()
            else:
                dialogue_list = extract_dialogue_from_docx(st.session_state.book_name, st.session_state.docx_path)
                st.session_state.quotes_lines = dialogue_list
                st.session_state.docx_only = True
                st.success("Quotes extracted from DOCX.")
                quotes_txt = "\n".join(dialogue_list)
                st.download_button("Download Extracted Quotes TXT", quotes_txt.encode("utf-8"),
                                   file_name=f"{st.session_state.book_name}-quotes.txt", mime="text/plain")
                if st.button("Restart", key="restart_docx"):
                    restart_app()
                if st.button("Continue", key="continue_docx"):
                    st.session_state.docx_only = False
                    st.session_state.unknown_index = 0
                    st.session_state.console_log = []
                    st.session_state.step = 2
                    auto_save()
                    st.rerun()
        else:
            pass
    else:
        docx_file = st.file_uploader("Upload DOCX File", type=["docx"])
        quotes_file = st.file_uploader("Upload Quotes TXT File (optional)", type=["txt"])
        speaker_colors_file = st.file_uploader("Upload Speaker Colors JSON (optional)", type=["json"])
        
        if st.button("Start Processing"):
            if docx_file is None:
                st.error("Please upload a DOCX file.")
            else:
                st.session_state.book_name = os.path.splitext(docx_file.name)[0]
                st.session_state.docx_bytes = docx_file.getvalue()
                with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_docx:
                    tmp_docx.write(st.session_state.docx_bytes)
                    st.session_state.docx_path = tmp_docx.name
                if quotes_file is not None:
                    quotes_text = quotes_file.read().decode("utf-8")
                    st.session_state.quotes_lines = quotes_text.splitlines(keepends=True)
                    st.session_state.docx_only = False
                else:
                    st.session_state.quotes_lines = None
                    st.session_state.docx_only = True
                if speaker_colors_file is not None:
                    raw = json.load(speaker_colors_file)
                    st.session_state.existing_speaker_colors = {normalize_speaker_name(k): v for k, v in raw.items()}
                    save_speaker_colors(st.session_state.existing_speaker_colors)
                else:
                    st.session_state.existing_speaker_colors = {}
                st.session_state.unknown_index = 0
                st.session_state.console_log = []
                if st.session_state.docx_only:
                    st.session_state.step = 1
                else:
                    st.session_state.step = 2
                auto_save()
                st.rerun()

# ========= STEP 2: Unknown Speaker Processing =========
elif st.session_state.step == 2:
    st.markdown("<h4>Step 2: Process Unknown Speakers</h4>", unsafe_allow_html=True)
    st.write("For each quote with speaker 'Unknown', type a replacement (or type 'skip', 'exit', or 'undo').")
    
    def get_next_unknown_line():
        quotes = st.session_state.get("quotes_lines")
        if quotes is None:
            return None, None, None
        pattern = re.compile(r"^(\s*\d+(?:[a-zA-Z]+)?\.\s+)([^:]+)(:.*)$")
        for i in range(st.session_state.unknown_index, len(quotes)):
            line = quotes[i]
            m = pattern.match(line)
            if m:
                prefix, speaker_raw, remainder = m.groups()
                if speaker_raw.strip() == "Unknown":
                    return i, prefix, remainder
        return None, None, None

    index, prefix, remainder = get_next_unknown_line()
    if index is None:
        st.write("No more unknown speakers found.")
        if st.button("Proceed to Color Assignment"):
            st.session_state.step = 3
            auto_save()
            st.rerun()
    else:
        dialogue = remainder.lstrip(": ").rstrip("\n")
        st.markdown("<hr style='margin: 2px 0;'>", unsafe_allow_html=True)
        def get_context_for_dialogue(dialogue):
            try:
                doc = docx.Document(st.session_state.docx_path)
            except Exception:
                return None
            normalized_dialogue = normalize_text(dialogue).lower()
            for idx, para in enumerate(doc.paragraphs):
                para_text = normalize_text(para.text)
                if normalized_dialogue in para_text.lower():
                    context = {}
                    if idx > 0:
                        context['previous'] = doc.paragraphs[idx-1].text
                    pattern = re.compile(re.escape(dialogue), re.IGNORECASE)
                    highlighted = pattern.sub(lambda m: f"<b>{m.group(0)}</b>", para.text)
                    context['current'] = highlighted
                    if idx+1 < len(doc.paragraphs):
                        context['next'] = doc.paragraphs[idx+1].text
                    return context
            return None

        context = get_context_for_dialogue(dialogue)
        if context:
            if "previous" in context:
                st.write(context["previous"])
            st.markdown(context["current"], unsafe_allow_html=True)
            if "next" in context:
                st.write(context["next"])
        else:
            st.write("No context found in DOCX for this quote.")
        st.markdown("<hr style='margin: 2px 0;'>", unsafe_allow_html=True)
        st.write(f"**Dialogue (Line {index+1}):** {dialogue}")
        
        def process_unknown_input():
            new_speaker = st.session_state.new_speaker_input.strip()
            if new_speaker.lower() == "exit":
                st.session_state.console_log.insert(0, "Exiting unknown speaker processing.")
                st.session_state.step = 3
            elif new_speaker.lower() == "skip":
                st.session_state.console_log.insert(0, f"Skipped line {index+1}.")
                st.session_state.unknown_index = index + 1
            elif new_speaker.lower() == "undo":
                if "last_update" in st.session_state:
                    last_index = st.session_state.last_update[0]
                    pattern = re.compile(r"^(\s*\d+(?:[a-zA-Z]+)?\.\s+)([^:]+)(:.*)$")
                    m = pattern.match(st.session_state.quotes_lines[last_index])
                    if m:
                        prefix_u, _, remainder_u = m.groups()
                        st.session_state.quotes_lines[last_index] = prefix_u + "Unknown" + remainder_u
                        st.session_state.unknown_index = last_index
                        st.session_state.console_log.insert(0, f"Reverted line {last_index+1} to Unknown.")
                    del st.session_state.last_update
                else:
                    st.session_state.console_log.insert(0, "Nothing to undo.")
            else:
                st.session_state.last_update = (index, st.session_state.quotes_lines[index])
                updated_speaker = smart_title(new_speaker)
                new_line = prefix + updated_speaker + remainder
                if not new_line.endswith("\n"):
                    new_line += "\n"
                st.session_state.quotes_lines[index] = new_line
                st.session_state.console_log.insert(0, f"Updated line {index+1} with speaker: {updated_speaker}")
                st.session_state.unknown_index = index + 1
            st.session_state.new_speaker_input = ""
            auto_save()
            st.session_state.setdefault("edit_counter", 0)
            st.session_state.edit_counter += 1
            if st.session_state.edit_counter % REFRESH_EVERY == 0:
                st.session_state.pending_refresh = True
        
        st.text_input("Enter speaker name (or 'skip'/'exit'/'undo'):", key="new_speaker_input", on_change=process_unknown_input)
        st.text_area("Console Log", "\n".join(st.session_state.console_log), height=150, label_visibility="collapsed")

# ========= STEP 3: Speaker Color Assignment =========
elif st.session_state.step == 3:
    st.markdown("<h4>Step 3: Speaker Color Assignment</h4>", unsafe_allow_html=True)
    st.write("Assign highlight colors for speakers that do not yet have an assigned color. You can also click 'Edit Speaker Colors' to review and change all assignments.")
    # Load the canonical speakers.
    with tempfile.NamedTemporaryFile(delete=False, suffix=".txt", mode="w+", encoding="utf-8") as tmp_quotes:
        tmp_quotes.write("".join(st.session_state.quotes_lines))
        tmp_quotes_path = tmp_quotes.name
    canonical_speakers, canonical_map = get_canonical_speakers(tmp_quotes_path)
    st.session_state.canonical_map = canonical_map
    # Load existing colors (or default to empty dict)
    existing_colors = st.session_state.get("existing_speaker_colors") or load_existing_colors() or {}
    
    # Determine which speakers need a new assignment
    speakers_to_assign = [
        sp for sp in canonical_speakers 
        if sp.lower() != "unknown" and (normalize_speaker_name(sp) not in existing_colors or existing_colors.get(normalize_speaker_name(sp), "none") == "none")
    ]
    
    if speakers_to_assign:
        st.write("Assign colors to the following speakers:")
        color_options = [color.title() for color in COLOR_PALETTE.keys()]
        updated_colors = {}
        for sp in speakers_to_assign:
            norm = normalize_speaker_name(sp)
            default_color = existing_colors.get(norm, "none")
            try:
                default_index = color_options.index(default_color.title())
            except ValueError:
                default_index = color_options.index("None")
            selected = st.selectbox(sp, options=color_options, index=default_index, key="new_"+norm)
            updated_colors[norm] = selected.lower()
        # Merge updated colors with any already assigned values.
        for norm, col in updated_colors.items():
            existing_colors[norm] = col
        # Build final dictionary using normalized keys.
        final_colors = {}
        for sp in canonical_speakers:
            norm = normalize_speaker_name(sp)
            if sp.lower() == "unknown":
                final_colors[norm] = "none"
            else:
                final_colors[norm] = existing_colors.get(norm, "none")
        st.session_state.speaker_colors = final_colors
        st.session_state.existing_speaker_colors = existing_colors.copy()
        save_speaker_colors(final_colors)
        st.success("Speaker colors updated.")
    else:
        st.write("All speakers already have assigned colors.")
    if os.path.exists(SAVED_COLORS_FILE):
        with open(SAVED_COLORS_FILE, "r", encoding="utf-8") as f:
            loaded_colors = json.load(f)
        st.session_state.speaker_colors = loaded_colors
        st.session_state.existing_speaker_colors = {normalize_speaker_name(k): v for k, v in loaded_colors.items()}
    else:
        st.session_state.speaker_colors = {}
        st.session_state.existing_speaker_colors = {}
    col1, col2 = st.columns(2)
    with col1:
        if st.button("Continue"):
            st.session_state.step = 4
            auto_save()
            st.rerun()
    with col2:
        if st.button("Edit Speaker Colors"):
            if os.path.exists(SAVED_COLORS_FILE):
                with open(SAVED_COLORS_FILE, "r", encoding="utf-8") as f:
                    loaded_colors = json.load(f)
                st.session_state.speaker_colors = loaded_colors
                st.session_state.existing_speaker_colors = {normalize_speaker_name(k): v for k, v in loaded_colors.items()}
            st.session_state.step = "edit_colors"
            auto_save()
            st.rerun()

# ========= EDIT COLORS: Full Speaker Color Assignment =========
elif st.session_state.step == "edit_colors":
    st.markdown("<h4>Edit Speaker Colors</h4>", unsafe_allow_html=True)
    st.write("Edit the assigned colors for all speakers (excluding 'Unknown'):")
    with tempfile.NamedTemporaryFile(delete=False, suffix=".txt", mode="w+", encoding="utf-8") as tmp_quotes:
        tmp_quotes.write("".join(st.session_state.quotes_lines))
        tmp_quotes_path = tmp_quotes.name
    canonical_speakers, canonical_map = get_canonical_speakers(tmp_quotes_path)
    st.session_state.canonical_map = canonical_map
    # Load current colors (or default to empty)
    existing_colors = st.session_state.get("speaker_colors") or load_existing_colors() or {}
    updated_colors = existing_colors.copy()
    color_options = [color.title() for color in COLOR_PALETTE.keys()]
    for sp in canonical_speakers:
        if sp.lower() == "unknown":
            continue
        norm = normalize_speaker_name(sp)
        default_color = existing_colors.get(norm, "none")
        try:
            default_index = color_options.index(default_color.title())
        except ValueError:
            default_index = color_options.index("None")
        selected = st.selectbox(sp, options=color_options, index=default_index, key="edit_"+norm)
        updated_colors[norm] = selected.lower()
    st.session_state.speaker_colors = updated_colors
    st.session_state.existing_speaker_colors = updated_colors.copy()
    save_speaker_colors(updated_colors)
    st.success("Speaker colors updated.")
    if st.button("Continue"):
        st.session_state.step = 4
        auto_save()
        st.rerun()

# ========= STEP 4: Final HTML Generation =========
elif st.session_state.step == 4:
    if "speaker_colors" not in st.session_state or st.session_state.speaker_colors is None:
        st.session_state.speaker_colors = load_existing_colors() or {}
    st.markdown("<h4>Step 4: Final HTML Generation</h4>", unsafe_allow_html=True)
    with tempfile.NamedTemporaryFile(delete=False, suffix=".txt", mode="w+", encoding="utf-8") as tmp_quotes:
        tmp_quotes.write("".join(st.session_state.quotes_lines))
        quotes_file_path = tmp_quotes.name
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_marker:
        marker_docx_path = tmp_marker.name
    create_marker_docx(st.session_state.docx_path, marker_docx_path)
    html = convert_docx_to_html_mammoth(marker_docx_path)
    os.remove(marker_docx_path)
    quotes_list = load_quotes(quotes_file_path, st.session_state.canonical_map)
    highlighted_html = highlight_dialogue_in_html(html, quotes_list, st.session_state.speaker_colors)
    final_html_body = apply_manual_indentation_with_markers(st.session_state.docx_path, highlighted_html)
    summary_html = generate_summary_html(quotes_list, list(st.session_state.canonical_map.values()), st.session_state.speaker_colors)
    ranking_html = generate_ranking_html(quotes_list, st.session_state.speaker_colors)
    final_html_body = summary_html + "\n<br><br><br>\n" + ranking_html + "\n" + final_html_body
    final_html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>{st.session_state.book_name}</title>
  <style>
    body {{
      font-family: Avenir, sans-serif;
      line-height: 2;
      max-width: 500px;
      margin: auto;
    }}
    span {{
      padding: 0;
    }}
    span.highlight {{
      background-color: var(--highlight-color, transparent);
      padding: 0.33em 0px;
      box-decoration-break: clone;
      -webkit-box-decoration-break: clone;
    }}
  </style>
</head>
<body>
{final_html_body}
</body>
</html>
"""
    final_html_path = os.path.join(tempfile.gettempdir(), f"{st.session_state.book_name}.html")
    with open(final_html_path, "w", encoding="utf-8") as f:
        f.write(final_html)
    st.success("Final HTML generated.")
    components.html(final_html, height=800, scrolling=True)
    with open(final_html_path, "rb") as f:
        html_bytes = f.read()
    st.download_button("Download HTML File", html_bytes,
                       file_name=f"{st.session_state.book_name}.html", mime="text/html")
    updated_colors = json.dumps(st.session_state.speaker_colors, indent=4, ensure_ascii=False).encode("utf-8")
    st.download_button("Download Updated Speaker Colors JSON", updated_colors,
                       file_name="speaker_colors.json", mime="application/json")
    updated_quotes = "".join(st.session_state.quotes_lines).encode("utf-8")
    st.download_button("Download Updated Quotes TXT", updated_quotes,
                       file_name=f"{st.session_state.book_name}-quotes.txt", mime="text/plain")
    if os.path.exists("unmatched_quotes.txt"):
        with open("unmatched_quotes.txt", "rb") as f:
            unmatched_bytes = f.read()
        st.download_button("Download Unmatched Quotes TXT", unmatched_bytes,
                           file_name="unmatched_quotes.txt", mime="text/plain")
    if st.button("Return to Step 2"):
        if "book_name" in st.session_state:
            quotes_filename = f"{st.session_state.book_name}-quotes.txt"
            if os.path.exists(quotes_filename):
                with open(quotes_filename, "r", encoding="utf-8") as f:
                    st.session_state.quotes_lines = f.read().splitlines(keepends=True)
        if os.path.exists("speaker_colors.json"):
            with open("speaker_colors.json", "r", encoding="utf-8") as f:
                colors = json.load(f)
            st.session_state.speaker_colors = colors
            st.session_state.existing_speaker_colors = {normalize_speaker_name(k): v for k, v in colors.items()}
        st.session_state.step = 2
        auto_save()
        st.rerun()
