import streamlit as st
import re
import os
import json
import tempfile
import mammoth
import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from bs4 import BeautifulSoup, NavigableString
from collections import Counter

# Import components for HTML embedding.
import streamlit.components.v1 as components

# ---------------------------
# Global Constants & Helper Functions
# ---------------------------
COLOR_PALETTE = {
    "dark grey": (67, 62, 63, 0.5, "black"),
    "burgundy": (134, 8, 0, 0.5, "black"),
    "red": (255, 10, 0, 0.5, "black"),
    "orange": (255, 113, 0, 0.5, "black"),
    "yellow": (255, 221, 0, 0.5, "black"),
    "dark yellow": (190, 174, 0, 0.5, "black"),
    "brown": (129, 100, 51, 0.5, "black"),
    "silver": (224, 224, 224, 0.5, "black"),
    "light green": (17, 255, 0, 0.5, "black"),
    "dark green": (9, 97, 0, 0.5, "black"),
    "turquoise": (0, 255, 206, 0.5, "black"),
    "light blue": (0, 237, 255, 0.5, "black"),
    "bright blue": (0, 32, 255, 0.5, "black"),
    "dark blue": (0, 32, 255, 0.5, "black"),
    "navy blue": (0, 25, 119, 0.5, "black"),
    "dark purple": (164, 49, 172, 0.6, "black"),
    "light purple": (196, 125, 255, 0.3, "black"),
    "bright pink": (255, 0, 207, 0.5, "black"),
    "light pink": (249, 212, 234, 0.4, "black"),
    "pale pink": (249, 212, 234, 0.4, "black"),
    "wine": (202, 78, 78, 0.5, "black"),
    "lime": (193, 227, 71, 0.5, "black"),
    "none": (134, 8, 0, 1.0, "rgb(134, 8, 0)")
}
SAVED_COLORS_FILE = "speaker_colors.json"
PROGRESS_FILE = "progress.json"  # file to store auto-saved progress

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
    return name.replace(".", "").lower().strip()

def smart_title(name):
    """
    Returns a normalized title for a speaker name.
    Each word is capitalized (first letter uppercase, rest lowercase),
    except that words in the exceptions (ps, pc, ds, di, dci) are forced to uppercase.
    Also, any letter immediately following an apostrophe is forced to lowercase.
    """
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
    result = re.sub(r"\'([A-Z])", lambda m: "'" + m.group(1).lower(), result)
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
        "quotes_lines": st.session_state.get("quotes_lines"),
        "speaker_colors": st.session_state.get("speaker_colors"),
        "unknown_index": st.session_state.get("unknown_index", 0),
        "console_log": st.session_state.get("console_log", []),
        "canonical_map": st.session_state.get("canonical_map")
    }
    with open(PROGRESS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4)

def auto_load():
    if os.path.exists(PROGRESS_FILE):
        with open(PROGRESS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        for key, value in data.items():
            st.session_state[key] = value

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
    st.write("Marker DOCX created.")

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
        color_choice = speaker_colors.get(speaker, "none")
        if speaker.lower() == "unknown":
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
        color_key = speaker_colors.get(sp, "none")
        if sp.lower() == "unknown":
            color_key = "none"
        rgba = COLOR_PALETTE.get(color_key, COLOR_PALETTE["none"])
        if color_key == "none":
            style = f"color: rgb({rgba[0]}, {rgba[1]}, {rgba[2]}); background-color: transparent;"
        else:
            style = f"color: {rgba[4]}; background-color: rgba({rgba[0]}, {rgba[1]}, {rgba[2]}, {rgba[3]});"
        lines.append(f'<p style="margin: 0; line-height: 1.2; padding: 5px 0;"><span class="highlight" style="{style}">{sp}</span> - {count} lines - {percentage}%</p>')
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
        color_key = speaker_colors.get(sp, "none")
        rgba = COLOR_PALETTE.get(color_key, COLOR_PALETTE["none"])
        if color_key == "none":
            style = f"color: rgb({rgba[0]}, {rgba[1]}, {rgba[2]}); background-color: transparent;"
        else:
            style = f"color: {rgba[4]}; background-color: rgba({rgba[0]}, {rgba[1]}, {rgba[2]}, {rgba[3]});"
        lines.append(f'<p style="margin: 0; line-height: 1.2; padding: 5px 0;"><span class="highlight" style="{style}">{sp}</span> - {count} lines - {percentage}%</p>')
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
                speakers.append(smart_title(speaker_raw))
    seen = set()
    canonical_speakers = []
    for s in speakers:
        norm = normalize_speaker_name(s)
        if norm not in seen:
            seen.add(norm)
            canonical_speakers.append(s)
    canonical_map = {normalize_speaker_name(s): s for s in canonical_speakers}
    st.write("Canonical speakers (after correction):", canonical_speakers)
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
            return json.load(f)
    return {}

def save_speaker_colors(speaker_colors):
    with open(SAVED_COLORS_FILE, "w", encoding="utf-8") as f:
        json.dump(speaker_colors, f, indent=4)

# ---------------------------
# Streamlit Multi-Step UI
# ---------------------------
if 'step' not in st.session_state:
    st.session_state.step = 1

# ========= STEP 1: Upload & Initialize =========
if st.session_state.step == 1:
    st.title("DOCX to HTML Converter with Dialogue Highlighting")
    
    # Check for saved progress and offer to load it.
    if os.path.exists(PROGRESS_FILE):
        st.info("Saved progress found. Click the button below to load your progress.")
        if st.button("Load Progress"):
            auto_load()
            st.success("Progress loaded.")
            st.rerun()
    
    st.write("Upload your DOCX and quotes text files. Optionally, upload an existing speaker_colors.json file.")
    docx_file = st.file_uploader("Upload DOCX File", type=["docx"])
    quotes_file = st.file_uploader("Upload Quotes TXT File", type=["txt"])
    speaker_colors_file = st.file_uploader("Upload Speaker Colors JSON (optional)", type=["json"])
    
    if st.button("Start Processing"):
        if docx_file is None or quotes_file is None:
            st.error("Please provide both the DOCX and Quotes files.")
        else:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_docx:
                tmp_docx.write(docx_file.getvalue())
                st.session_state.docx_path = tmp_docx.name
            st.session_state.book_name = os.path.splitext(docx_file.name)[0]
            quotes_text = quotes_file.read().decode("utf-8")
            st.session_state.quotes_lines = quotes_text.splitlines(keepends=True)
            if speaker_colors_file is not None:
                st.session_state.existing_speaker_colors = json.load(speaker_colors_file)
            else:
                st.session_state.existing_speaker_colors = {}
            st.session_state.unknown_index = 0
            st.session_state.console_log = []
            st.session_state.step = 2
            auto_save()  # Save progress after initialization.
            st.rerun()

# ========= STEP 2: Unknown Speaker Processing =========
elif st.session_state.step == 2:
    st.title("Step 2: Process Unknown Speakers")
    st.write("For each quote with speaker 'Unknown', type a replacement (or type 'skip'/'exit').")
    
    def get_next_unknown_line():
        pattern = re.compile(r"^(\s*\d+(?:[a-zA-Z]+)?\.\s+)([^:]+)(:.*)$")
        for i in range(st.session_state.unknown_index, len(st.session_state.quotes_lines)):
            line = st.session_state.quotes_lines[i]
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
        st.write(f"**Line {index+1}:**")
        st.write("**Dialogue:**", dialogue)
        
        def get_context_for_dialogue(dialogue):
            try:
                doc = docx.Document(st.session_state.docx_path)
            except Exception:
                return None
            normalized_dialogue = normalize_text(dialogue).lower()
            for idx, para in enumerate(doc.paragraphs):
                para_text = normalize_text(para.text).lower()
                if normalized_dialogue in para_text:
                    context = {}
                    if idx > 0:
                        context['previous'] = doc.paragraphs[idx-1].text
                    context['current'] = doc.paragraphs[idx].text
                    if idx+1 < len(doc.paragraphs):
                        context['next'] = doc.paragraphs[idx+1].text
                    return context
            return None

        context = get_context_for_dialogue(dialogue)
        if context:
            st.write("**Context:**")
            if "previous" in context:
                st.write("*Previous Paragraph:*", context["previous"])
            st.write("*Current Paragraph:*", context["current"])
            if "next" in context:
                st.write("*Next Paragraph:*", context["next"])
        else:
            st.write("No context found in DOCX for this quote.")
        
        def process_unknown_input():
            new_speaker = st.session_state.new_speaker_input.strip()
            if new_speaker.lower() == "exit":
                st.session_state.console_log.append("Exiting unknown speaker processing.")
                st.session_state.step = 3
            elif new_speaker.lower() == "skip":
                st.session_state.console_log.append(f"Skipped line {index+1}.")
                st.session_state.unknown_index = index + 1
            else:
                updated_speaker = smart_title(new_speaker)
                new_line = prefix + updated_speaker + remainder
                if not new_line.endswith("\n"):
                    new_line += "\n"
                st.session_state.quotes_lines[index] = new_line
                st.session_state.console_log.append(f"Updated line {index+1} with speaker: {updated_speaker}")
                st.session_state.unknown_index = index + 1
            st.session_state.new_speaker_input = ""
            auto_save()
        
        st.text_input("Enter speaker name (or 'skip'/'exit'):", key="new_speaker_input", on_change=process_unknown_input)
        st.text_area("Console Log", "\n".join(st.session_state.console_log), height=150, label_visibility="collapsed")

# ========= STEP 3: Speaker Color Assignment =========
elif st.session_state.step == 3:
    st.title("Step 3: Speaker Color Assignment")
    st.write("The app extracts canonical speakers from the quotes file.")
    with tempfile.NamedTemporaryFile(delete=False, suffix=".txt", mode="w+", encoding="utf-8") as tmp_quotes:
        tmp_quotes.write("".join(st.session_state.quotes_lines))
        tmp_quotes_path = tmp_quotes.name
    canonical_speakers, canonical_map = get_canonical_speakers(tmp_quotes_path)
    st.session_state.canonical_map = canonical_map
    st.write("Canonical Speakers:", canonical_speakers)
    existing_colors = st.session_state.existing_speaker_colors if "existing_speaker_colors" in st.session_state else load_existing_colors()
    st.write("Select a highlight color for each speaker (for 'Unknown', it is fixed to 'none').")
    speaker_colors = {}
    with st.form("color_assignment_form"):
        for sp in canonical_speakers:
            if sp.lower() == "unknown":
                speaker_colors[sp] = "none"
                st.write(f"{sp}: none")
            else:
                default_color = existing_colors.get(sp, "none")
                speaker_colors[sp] = st.selectbox(
                    f"Color for {sp}",
                    options=list(COLOR_PALETTE.keys()),
                    index=list(COLOR_PALETTE.keys()).index(default_color),
                    key=sp
                )
        submitted = st.form_submit_button("Submit Colors")
        if submitted:
            st.session_state.speaker_colors = speaker_colors
            save_speaker_colors(speaker_colors)
            st.session_state.step = 4
            auto_save()
            st.rerun()

# ========= STEP 4: Final HTML Generation =========
elif st.session_state.step == 4:
    st.title("Step 4: Final HTML Generation")
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
      max-width: 800px;
      margin: auto;
    }}
    span {{
      padding: 0;
    }}
    span.highlight {{
      background-color: var(--highlight-color, transparent);
      padding: 0.15em 0px;
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
    st.markdown("### Final HTML Preview")
    components.html(final_html, height=800, scrolling=True)
    with open(final_html_path, "rb") as f:
        html_bytes = f.read()
    st.download_button("Download HTML File", html_bytes, file_name=f"{st.session_state.book_name}.html", mime="text/html")
    updated_colors = json.dumps(st.session_state.speaker_colors, indent=4).encode("utf-8")
    st.download_button("Download Updated Speaker Colors JSON", updated_colors, file_name=f"{st.session_state.book_name}-speaker_colors.json", mime="application/json")
    updated_quotes = "".join(st.session_state.quotes_lines).encode("utf-8")
    st.download_button("Download Updated Quotes TXT", updated_quotes, file_name=f"{st.session_state.book_name}-quotes.txt", mime="text/plain")