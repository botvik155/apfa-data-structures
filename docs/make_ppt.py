#!/usr/bin/env python3
"""
Builds the 8-slide presentation for DS Project 2 (Hostel Student Management System).
Run:  python3 docs/make_ppt.py     ->  docs/DS_Project_2_Presentation.pptx
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

# ---------- theme ----------
NAVY   = RGBColor(0x12, 0x35, 0x6F)
BLUE   = RGBColor(0x2F, 0x6F, 0xED)
GREEN  = RGBColor(0x16, 0xA3, 0x4A)
RED    = RGBColor(0xDC, 0x26, 0x26)
DARK   = RGBColor(0x1A, 0x1A, 0x1A)
GREY   = RGBColor(0x55, 0x5F, 0x6D)
LIGHT  = RGBColor(0xF3, 0xF6, 0xFB)
CODEBG = RGBColor(0x1E, 0x24, 0x30)
CODEFG = RGBColor(0xE6, 0xE9, 0xEF)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
BORDER = RGBColor(0xD3, 0xDA, 0xE5)

BODY = "Calibri"
MONO = "Consolas"

W, H = Inches(13.333), Inches(7.5)

prs = Presentation()
prs.slide_width, prs.slide_height = W, H


def blank():
    return prs.slides.add_slide(prs.slide_layouts[6])


def box(slide, x, y, w, h, fill=None, line=None, lw=1.25):
    from pptx.enum.shapes import MSO_SHAPE
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    if fill is None:
        sh.fill.background()
    else:
        sh.fill.solid()
        sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = Pt(lw)
    sh.shadow.inherit = False
    return sh


def text(slide, x, y, w, h, runs, size=16, color=DARK, font=BODY,
         align=PP_ALIGN.LEFT, bold=False, space=6, line=None, anchor=MSO_ANCHOR.TOP):
    """runs = string, or list of paragraphs; a paragraph is a string or a list of
       (text, bold, color) tuples."""
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0

    paras = [runs] if isinstance(runs, str) else runs
    for i, p in enumerate(paras):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = align
        para.space_after = Pt(space)
        if line:
            para.line_spacing = line
        pieces = [(p, bold, color)] if isinstance(p, str) else p
        for t, b, c in pieces:
            r = para.add_run()
            r.text = t
            r.font.size = Pt(size)
            r.font.bold = b
            r.font.color.rgb = c
            r.font.name = font
    return tb


def header(slide, title, kicker=None):
    box(slide, 0, 0, W, Inches(1.0), fill=NAVY)
    text(slide, Inches(0.55), Inches(0.24), Inches(11.0), Inches(0.55),
         title, size=27, color=WHITE, bold=True)
    if kicker:
        text(slide, Inches(11.3), Inches(0.34), Inches(1.6), Inches(0.4),
             kicker, size=13, color=RGBColor(0x9F, 0xB6, 0xE8), bold=True,
             align=PP_ALIGN.RIGHT)


def code(slide, x, y, w, lines, size=13.5, accent=BLUE):
    h = Inches(0.30 + 0.245 * len(lines))
    box(slide, x, y, w, h, fill=CODEBG)
    box(slide, x, y, Inches(0.055), h, fill=accent)
    tb = slide.shapes.add_textbox(x + Inches(0.22), y + Inches(0.14),
                                  w - Inches(0.34), h - Inches(0.2))
    tf = tb.text_frame
    tf.word_wrap = False
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, ln in enumerate(lines):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.space_after = Pt(0)
        para.line_spacing = 1.12
        # anything after '//' is a comment -> green
        main, sep, com = ln.partition("//")
        r = para.add_run(); r.text = main
        r.font.size = Pt(size); r.font.name = MONO; r.font.color.rgb = CODEFG
        if sep:
            r2 = para.add_run(); r2.text = sep + com
            r2.font.size = Pt(size); r2.font.name = MONO
            r2.font.color.rgb = RGBColor(0x7E, 0xC9, 0x9B)
    return y + h


def bullets(slide, x, y, w, items, size=16.5, gap=10, marker=BLUE):
    """items = list of (bold_lead, rest) or plain strings."""
    cur = y
    for it in items:
        box(slide, x, cur + Inches(0.085), Inches(0.075), Inches(0.075), fill=marker)
        if isinstance(it, str):
            runs = [(it, False, DARK)]
        else:
            runs = [(it[0], True, NAVY), (it[1], False, DARK)]
        tb = text(slide, x + Inches(0.26), cur, w - Inches(0.26), Inches(0.4),
                  [runs], size=size, line=1.25)
        # estimate height for the next bullet
        chars = sum(len(t) for t, _, _ in runs)
        per_line = int(w.inches * 100 / (size * 0.062) / 10)
        n = max(1, -(-chars // max(20, per_line)))
        cur += Inches(0.02) + Emu(int(n * size * 1.42 * 12700)) + Emu(int(gap * 12700))
    return cur


def table(slide, x, y, w, rows, widths, size=13.5, head=True, rh=0.38):
    """rows[0] = header row."""
    n, m = len(rows), len(rows[0])
    tbl = slide.shapes.add_table(n, m, x, y, w, Inches(rh * n)).table
    total = sum(widths)
    for j, ww in enumerate(widths):
        tbl.columns[j].width = Emu(int(w * ww / total))
    for i, row in enumerate(rows):
        tbl.rows[i].height = Inches(rh if i else rh)
        for j, val in enumerate(row):
            cell = tbl.cell(i, j)
            cell.text = ""
            cell.margin_left = cell.margin_right = Inches(0.09)
            cell.margin_top = cell.margin_bottom = Inches(0.03)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.fill.solid()
            if i == 0 and head:
                cell.fill.fore_color.rgb = NAVY
            else:
                cell.fill.fore_color.rgb = WHITE if i % 2 else LIGHT
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER if (j and len(val) < 14) else PP_ALIGN.LEFT
            r = p.add_run(); r.text = val
            r.font.size = Pt(size)
            r.font.name = MONO if val.startswith(("push", "pop", "top_", "count_", "O(")) else BODY
            r.font.bold = (i == 0 and head)
            r.font.color.rgb = WHITE if (i == 0 and head) else DARK
    return tbl


def footer(slide, n):
    text(slide, Inches(0.55), Inches(7.02), Inches(8.0), Inches(0.3),
         "Hostel Student Management System  |  Array, Stack, String  |  C++",
         size=11, color=GREY)
    text(slide, Inches(12.2), Inches(7.02), Inches(0.6), Inches(0.3),
         str(n), size=11, color=GREY, align=PP_ALIGN.RIGHT)


# =====================================================================
# 1 — TITLE
# =====================================================================
s = blank()
box(s, 0, 0, W, H, fill=NAVY)
box(s, 0, Inches(4.62), W, Inches(0.055), fill=BLUE)
text(s, Inches(1.0), Inches(2.05), Inches(11.5), Inches(1.0),
     "Hostel Student Management System", size=41, color=WHITE, bold=True)
text(s, Inches(1.0), Inches(3.15), Inches(11.3), Inches(0.6),
     [[("A Data Structures mini project in ", False, RGBColor(0xC3, 0xD2, 0xF0)),
       ("C++", True, WHITE)]], size=22)
text(s, Inches(1.0), Inches(3.72), Inches(11.3), Inches(0.5),
     [[("Arrays", True, WHITE), ("   •   ", False, BLUE),
       ("Stacks", True, WHITE), ("   •   ", False, BLUE),
       ("Strings", True, WHITE)]], size=21)
text(s, Inches(1.0), Inches(5.05), Inches(11.3), Inches(1.2),
     ["DS Project 2",
      "Console menu  +  web interface (HTML & CSS) on localhost",
      "No external libraries — every data structure written by hand"],
     size=15, color=RGBColor(0xA9, 0xBE, 0xE4), space=4)

# =====================================================================
# 2 — THE PROJECT AND WHY THESE STRUCTURES
# =====================================================================
s = blank(); header(s, "The Project and Its Data Structures", "01")
text(s, Inches(0.55), Inches(1.22), Inches(12.2), Inches(0.5),
     [[("Goal: ", True, NAVY),
       ("manage hostel student records — add, display, search, delete, undo a "
        "delete and sort by name — while implementing the data structures ", False, DARK),
       ("by hand", True, NAVY),
       (" instead of using ready-made containers.", False, DARK)]], size=16.5, line=1.22)

table(s, Inches(0.55), Inches(2.15), Inches(12.2), [
    ["Structure", "Where it is used", "Why this one was chosen"],
    ["ARRAY",
     "students[100] holds every record, with count_ tracking how many are in use",
     "Records need indexed access and ordered display, and the maximum is known in advance"],
    ["STACK",
     "Deleted records are pushed here; Undo pops the top one back into the array",
     "Undo must reverse the most recent action first — that is exactly LIFO"],
    ["STRING",
     "Name and course fields; case-insensitive search and alphabetical sorting",
     "Text needs character-level comparison, which std::string supports directly"],
], [1.6, 5.0, 5.6], size=14.5, rh=0.9)

box(s, Inches(0.55), Inches(6.0), Inches(12.2), Inches(0.88),
    fill=RGBColor(0xEC, 0xFD, 0xF3), line=GREEN)
text(s, Inches(0.85), Inches(6.18), Inches(11.6), Inches(0.6),
     [[("The key point: ", True, RGBColor(0x0F, 0x7A, 0x38)),
       ("the stack is not decoration. Undo has to reverse the ", False, DARK),
       ("most recent", True, DARK),
       (" deletion first. A queue would be wrong — it would undo the ", False, DARK),
       ("oldest", True, DARK), (" one first.", False, DARK)]], size=16, line=1.2)
footer(s, 2)

# =====================================================================
# 3 — THE ARRAY
# =====================================================================
s = blank(); header(s, "The Array — Storing the Records", "02")
code(s, Inches(0.55), Inches(1.3), Inches(5.85), [
    "Student students[MAX];   // MAX = 100",
    "int count_ = 0;          // how many are used",
    "",
    "bool addStudent(int roll, string name,",
    "                string course, int room) {",
    "  if (count_ >= MAX)   return false;  // full",
    "  if (findStudent(roll) != -1)",
    "      return false;    // duplicate roll",
    "",
    "  students[count_].roll   = roll;",
    "  students[count_].name   = name;",
    "  students[count_].course = course;",
    "  students[count_].room   = room;",
    "  count_++;",
    "  return true;",
    "}",
], size=13.5)

text(s, Inches(6.85), Inches(1.3), Inches(5.9), Inches(0.4),
     "HOW IT WORKS", size=13, color=BLUE, bold=True)
bullets(s, Inches(6.85), Inches(1.78), Inches(5.9), [
    ("Insert ", "puts the record at index count_ and increments it — O(1), no shifting needed"),
    ("count_ ", "is the logical size, kept separate from the physical capacity of 100"),
    ("Two checks first: ", "the array must not be full, and the roll number must not already exist"),
    ("Search ", "walks index 0 to count_-1 — a linear search, O(n)"),
    ("findStudent() returns the index", ", not the record, because deletion needs to know where it is"),
], size=15.5, gap=9)

box(s, Inches(6.85), Inches(5.5), Inches(5.9), Inches(0.95), fill=LIGHT, line=BORDER)
code(s, Inches(7.05), Inches(5.65), Inches(5.5), [
    "for (int i = 0; i < count_; i++)      // linear search",
    "    if (students[i].roll == roll) return i;",
], size=12.5)
footer(s, 3)

# =====================================================================
# 4 — THE STACK
# =====================================================================
s = blank(); header(s, "The Stack — Remembering Deletions", "03")
code(s, Inches(0.55), Inches(1.3), Inches(5.85), [
    "Student stack_[MAX];   // the stack",
    "int top_ = -1;         // -1 means empty",
    "",
    "void push(Student s) {",
    "    if (top_ < MAX - 1) {   // overflow check",
    "        top_++;",
    "        stack_[top_] = s;",
    "    }",
    "}",
    "",
    "Student pop() {",
    "    Student s = stack_[top_];",
    "    top_--;",
    "    return s;",
    "}",
], size=13.5, accent=GREEN)

text(s, Inches(6.85), Inches(1.3), Inches(5.9), Inches(0.4),
     "ARRAY IMPLEMENTATION OF A STACK", size=13, color=BLUE, bold=True)
table(s, Inches(6.85), Inches(1.75), Inches(5.9), [
    ["Operation", "What it does", "Time"],
    ["push(s)", "top_++ then store at that index", "O(1)"],
    ["pop()", "read at top_ then top_--", "O(1)"],
    ["stackEmpty()", "true when top_ == -1", "O(1)"],
], [1.5, 3.3, 0.9], size=13.5, rh=0.46)

bullets(s, Inches(6.85), Inches(3.75), Inches(5.9), [
    ("Overflow ", "= pushing onto a full stack, blocked by top_ < MAX-1"),
    ("Underflow ", "= popping an empty stack, blocked by top_ == -1"),
    ("Both checked before ", "the array is touched, so memory outside it is never accessed"),
    ("top_ starts at -1 ", "because an empty stack has no valid top index; size is always top_ + 1"),
], size=15.5, gap=9)
footer(s, 4)

# =====================================================================
# 5 — STRINGS
# =====================================================================
s = blank(); header(s, "Strings — Searching and Sorting", "04")

text(s, Inches(0.55), Inches(1.3), Inches(5.85), Inches(0.4),
     "CASE-INSENSITIVE SEARCH", size=13, color=BLUE, bold=True)
code(s, Inches(0.55), Inches(1.75), Inches(5.85), [
    "string toLower(string s) {",
    "  for (int i = 0; i < s.size(); i++)",
    "    if (s[i] >= 'A' && s[i] <= 'Z')",
    "        s[i] = s[i] + 32;   // 'A'=65 'a'=97",
    "  return s;",
    "}",
    "",
    "bool contains(string text, string key) {",
    "  return toLower(text).find(toLower(key))",
    "         != string::npos;",
    "}",
], size=13.5)
box(s, Inches(0.55), Inches(4.95), Inches(5.85), Inches(0.8), fill=LIGHT, line=BORDER)
text(s, Inches(0.78), Inches(5.13), Inches(5.5), Inches(0.55),
     [[("This is why typing ", False, DARK), ("bca", True, NAVY),
       (" also finds ", False, DARK), ("BCA", True, NAVY),
       (". The stored data is never changed — only the copies used for comparison.", False, DARK)]],
     size=14, line=1.2)

text(s, Inches(6.9), Inches(1.3), Inches(5.85), Inches(0.4),
     "SORTING BY NAME (BUBBLE SORT)", size=13, color=BLUE, bold=True)
code(s, Inches(6.9), Inches(1.75), Inches(5.85), [
    "void sortByName() {",
    " for (int i = 0; i < count_-1; i++)",
    "  for (int j = 0; j < count_-1-i; j++)",
    "   if (toLower(students[j].name) >",
    "       toLower(students[j+1].name)) {",
    "      Student t     = students[j];",
    "      students[j]   = students[j+1];",
    "      students[j+1] = t;      // swap",
    "   }",
    "}",
], size=13.5)
box(s, Inches(6.9), Inches(4.72), Inches(5.85), Inches(1.03), fill=LIGHT, line=BORDER)
text(s, Inches(7.13), Inches(4.9), Inches(5.5), Inches(0.75),
     [[("The ", False, DARK), (">", True, NAVY),
       (" operator on std::string compares ", False, DARK),
       ("lexicographically", True, NAVY),
       (" — character by character, like a dictionary. Nested loops make this ", False, DARK),
       ("O(n²)", True, NAVY), (".", False, DARK)]], size=14, line=1.2)

text(s, Inches(0.55), Inches(6.05), Inches(12.2), Inches(0.5),
     [[("One search box covers all three fields: it also matches the roll number "
        "typed as text, using ", False, DARK),
       ("to_string()", True, NAVY), (".", False, DARK)]], size=15, color=GREY)
footer(s, 5)

# =====================================================================
# 6 — DELETE + UNDO  (the key slide)
# =====================================================================
s = blank(); header(s, "Deletion and Undo — Both Structures Together", "05")

code(s, Inches(0.55), Inches(1.28), Inches(6.4), [
    "bool deleteStudent(int roll) {",
    "  int i = findStudent(roll);",
    "  if (i == -1) return false;   // not found",
    "",
    "  push(students[i]);      // STACK: save first",
    "",
    "  for (int j = i; j < count_-1; j++)",
    "      students[j] = students[j+1];  // shift",
    "  count_--;",
    "  return true;",
    "}",
], size=13.5, accent=RED)

box(s, Inches(0.55), Inches(4.28), Inches(6.4), Inches(2.55), fill=CODEBG)
tb = s.shapes.add_textbox(Inches(0.75), Inches(4.44), Inches(6.05), Inches(2.25))
tf = tb.text_frame; tf.word_wrap = False
tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
for i, (ln, col) in enumerate([
    ("DELETING ROLL 102", RGBColor(0x9F, 0xB6, 0xE8)),
    ("", CODEFG),
    ("before:  [101][102][103][104]   count_ = 4", CODEFG),
    ("                |", RGBColor(0x8A, 0x93, 0xA5)),
    ("         push a copy onto the stack", RGBColor(0x7E, 0xC9, 0x9B)),
    ("shift:   [101][103][104][104]   move each left", CODEFG),
    ("after:   [101][103][104]        count_ = 3", CODEFG),
    ("", CODEFG),
    ("stack:   | 102 |  <-- top_", RGBColor(0xF0, 0xB4, 0x5C)),
]):
    p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
    p.space_after = Pt(1); p.line_spacing = 1.16
    r = p.add_run(); r.text = ln
    r.font.size = Pt(13); r.font.name = MONO; r.font.color.rgb = col
    r.font.bold = (i == 0)

code(s, Inches(7.35), Inches(1.28), Inches(5.4), [
    "bool undoDelete() {",
    "  if (stackEmpty()) return false;",
    "  Student s = pop();   // STACK: top one",
    "  return addStudent(s.roll, s.name,",
    "                    s.course, s.room);",
    "}",
], size=13.5, accent=GREEN)

box(s, Inches(7.35), Inches(3.15), Inches(5.4), Inches(2.45), fill=LIGHT, line=BORDER)
tb = s.shapes.add_textbox(Inches(7.58), Inches(3.32), Inches(5.0), Inches(2.15))
tf = tb.text_frame; tf.word_wrap = False
tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
for i, (ln, col, bd) in enumerate([
    ("LIFO IN ACTION", BLUE, True),
    ("delete 101, 102, 103", NAVY, False),
    ("stack:  | 101 | 102 | 103 |  <- top", NAVY, False),
    ("", DARK, False),
    ("undo -> 103 back   (newest first)", RGBColor(0x0F, 0x7A, 0x38), False),
    ("undo -> 102 back", RGBColor(0x0F, 0x7A, 0x38), False),
    ("undo -> 101 back", RGBColor(0x0F, 0x7A, 0x38), False),
    ("undo -> empty, nothing happens", GREY, False),
]):
    p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
    p.space_after = Pt(2); p.line_spacing = 1.14
    r = p.add_run(); r.text = ln
    r.font.size = Pt(13); r.font.name = MONO; r.font.color.rgb = col; r.font.bold = bd

box(s, Inches(7.35), Inches(5.78), Inches(5.4), Inches(1.05),
    fill=RGBColor(0xFF, 0xF7, 0xE6), line=RGBColor(0xE0, 0xA8, 0x00))
text(s, Inches(7.58), Inches(5.95), Inches(4.95), Inches(0.75),
     [[("The order matters. ", True, RGBColor(0x8A, 0x5A, 0x00)),
       ("The record is pushed ", False, DARK), ("before", True, DARK),
       (" the shift overwrites it — otherwise undo would restore the wrong student.", False, DARK)]],
     size=14, line=1.2)
footer(s, 6)

# =====================================================================
# 7 — RUNNING IT + THE WEB UI
# =====================================================================
s = blank(); header(s, "Running It, and How the Web UI Works", "06")

text(s, Inches(0.55), Inches(1.25), Inches(6.0), Inches(0.35),
     "WEB VERSION  (use this for the demo)", size=13, color=BLUE, bold=True)
code(s, Inches(0.55), Inches(1.66), Inches(6.0), [
    "$ cd ds_project_2",
    "$ make web",
    "Server running -> http://localhost:8080",
], size=14, accent=GREEN)
text(s, Inches(0.55), Inches(2.95), Inches(6.0), Inches(0.35),
     "CONSOLE VERSION", size=13, color=BLUE, bold=True)
code(s, Inches(0.55), Inches(3.36), Inches(6.0), [
    "$ make cli",
    "1. Add student      5. Undo delete (stack)",
    "2. Show all         6. Sort by name",
    "3. Search student   7. Exit",
    "4. Delete student",
], size=14, accent=GREEN)
box(s, Inches(0.55), Inches(5.12), Inches(6.0), Inches(1.1),
    fill=RGBColor(0xFF, 0xF7, 0xE6), line=RGBColor(0xE0, 0xA8, 0x00))
text(s, Inches(0.78), Inches(5.3), Inches(5.55), Inches(0.8),
     [[("Run it from inside the project folder. ", True, RGBColor(0x8A, 0x5A, 0x00)),
       ("The server reads web/index.html by relative path, so starting it "
        "elsewhere gives a blank page.", False, DARK)]], size=14, line=1.2)
text(s, Inches(0.55), Inches(6.38), Inches(6.0), Inches(0.4),
     "Needs only g++ and a browser — nothing to install.",
     size=14, color=GREY)

text(s, Inches(6.9), Inches(1.25), Inches(5.85), Inches(0.5),
     [[("The UI is ", False, DARK), ("HTML and CSS only", True, NAVY),
       (" — no JavaScript. Every button is a normal form or link.", False, DARK)]],
     size=15, line=1.2)
box(s, Inches(6.9), Inches(2.05), Inches(5.85), Inches(3.1), fill=CODEBG)
tb = s.shapes.add_textbox(Inches(7.1), Inches(2.2), Inches(5.5), Inches(2.8))
tf = tb.text_frame; tf.word_wrap = False
tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
for i, (ln, col) in enumerate([
    ("browser              C++ server", RGBColor(0x9F, 0xB6, 0xE8)),
    ("GET /       ---->  read index.html", CODEFG),
    ("            <----  loop ARRAY -> rows", RGBColor(0x7E, 0xC9, 0x9B)),
    ("                   loop STACK -> list", RGBColor(0x7E, 0xC9, 0x9B)),
    ("                   fill {{ROWS}} etc.", RGBColor(0x7E, 0xC9, 0x9B)),
    ("", CODEFG),
    ("POST /add   ---->  addStudent()", CODEFG),
    ("GET /delete ---->  shift + push", CODEFG),
    ("GET /undo   ---->  pop", CODEFG),
    ("GET /sort   ---->  string compare", CODEFG),
    ("GET /?q=bca ---->  string search", CODEFG),
]):
    p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
    p.space_after = Pt(1); p.line_spacing = 1.16
    r = p.add_run(); r.text = ln
    r.font.size = Pt(13); r.font.name = MONO; r.font.color.rgb = col
box(s, Inches(6.9), Inches(5.3), Inches(5.85), Inches(1.5),
    fill=RGBColor(0xEC, 0xFD, 0xF3), line=GREEN)
text(s, Inches(7.13), Inches(5.48), Inches(5.4), Inches(1.2),
     [[("Good for the demo: ", True, RGBColor(0x0F, 0x7A, 0x38)),
       ("the page shows the stack itself, top marked ", False, DARK),
       ("TOP →", True, NAVY),
       (". Delete a student and it appears on the stack; press Undo and it "
        "disappears. The structure is visible on screen.", False, DARK)]],
     size=14, line=1.2)
footer(s, 7)

# =====================================================================
# 8 — RESULTS AND CONCLUSION
# =====================================================================
s = blank(); header(s, "Results and Conclusion", "07")

text(s, Inches(0.55), Inches(1.25), Inches(6.0), Inches(0.35),
     "TIME COMPLEXITY", size=13, color=BLUE, bold=True)
table(s, Inches(0.55), Inches(1.68), Inches(6.0), [
    ["Operation", "Structures", "Worst"],
    ["Add student", "Array", "O(n)"],
    ["Search", "Array + String", "O(n)"],
    ["Delete student", "Array + Stack", "O(n)"],
    ["Push / Pop", "Stack", "O(1)"],
    ["Undo delete", "Stack + Array", "O(n)"],
    ["Sort by name", "Array + String", "O(n²)"],
], [2.3, 2.6, 1.1], size=13.5, rh=0.4)
text(s, Inches(0.55), Inches(4.58), Inches(6.0), Inches(0.6),
     [[("Space: ", True, NAVY),
       ("O(n) — one array of 100 records plus one stack of 100, allocated once. "
        "Push and pop are the only O(1) worst case operations.", False, DARK)]],
     size=14, line=1.2)

text(s, Inches(6.9), Inches(1.25), Inches(5.85), Inches(0.35),
     "TESTING", size=13, color=BLUE, bold=True)
box(s, Inches(6.9), Inches(1.68), Inches(5.85), Inches(0.62), fill=RGBColor(0xEC, 0xFD, 0xF3), line=GREEN)
text(s, Inches(7.13), Inches(1.82), Inches(5.4), Inches(0.4),
     [[("12 test cases, all passing", True, RGBColor(0x0F, 0x7A, 0x38)),
       ("  — in both versions", False, DARK)]], size=15)
bullets(s, Inches(6.9), Inches(2.5), Inches(5.85), [
    ("Duplicate roll numbers ", "are rejected, array unchanged"),
    ("Deleting from the middle ", "shifts later records left, no gap"),
    ("The deleted record ", "appears on top of the stack"),
    ("Two deletes, two undos ", "restore in reverse order (LIFO)"),
    ("Undo on an empty stack ", "is handled safely, no crash"),
], size=14.5, gap=7)

box(s, Inches(0.55), Inches(5.35), Inches(12.2), Inches(1.5), fill=NAVY)
text(s, Inches(0.85), Inches(5.52), Inches(11.6), Inches(0.4),
     "CONCLUSION", size=12.5, color=RGBColor(0x9F, 0xB6, 0xE8), bold=True)
text(s, Inches(0.85), Inches(5.85), Inches(11.6), Inches(0.9),
     [[("Array", True, WHITE), (" for indexed storage,  ", False, RGBColor(0xC3, 0xD2, 0xF0)),
       ("Stack", True, WHITE), (" for undo because undo is inherently LIFO,  ", False, RGBColor(0xC3, 0xD2, 0xF0)),
       ("String", True, WHITE), (" for case-insensitive search and sorting.", False, RGBColor(0xC3, 0xD2, 0xF0))],
      [("All three written from first principles, in one header — ", False, RGBColor(0xC3, 0xD2, 0xF0)),
       ("src/hostel.h", True, WHITE),
       (" — ready to reuse in a larger project. Next step: a queue for the room waiting list.",
        False, RGBColor(0xC3, 0xD2, 0xF0))]],
     size=15.5, line=1.25, space=4)
footer(s, 8)

prs.save("docs/DS_Project_2_Presentation.pptx")
print("saved docs/DS_Project_2_Presentation.pptx with %d slides" % len(prs.slides._sldIdLst))
