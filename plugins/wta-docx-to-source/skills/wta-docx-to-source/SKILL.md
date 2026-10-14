---
name: "wta-docx-to-source"
description: Convert a WiseTech Academy content block Word script (.docx, the LD/DIY script template) into the page source code for the DIY editor — a single .html file using the WTA widget tags (wta-scroll-animator, wta-hotspot, wta-text-reveal, wta-expander, wta-html-import, wta-question, alert callouts, case-study callouts, glossary links, CDN images). Trigger when a learning designer uploads a script .docx and asks to "convert to source code", "transfer the script to code", "make the source code", "build the DIY source", "code this page", or "turn this docx into HTML for the LMS". Do NOT use for building standalone HTML interactions (use lxd-coder / content-block-html-imports) or for writing script content.
---

# WTA script (.docx) → DIY source code

Turns an LD Word script into the HTML source a DIY page needs, using the same
widget code as the production pages. Output is **one `.html` file per script**,
ready to paste into the DIY editor's source-code view. No preview page, no extra report file.

## Steps

1. **Find the uploaded .docx file(s)** in the uploads folder. You can process several scripts in one run.
   File names can contain non-breaking spaces; if a path fails, copy the file with a glob first
   (for example `cp /path/to/uploads/*M3P06* /tmp/script.docx`).

1b. **Mark the glossary items.** Green-highlighted text in learner-facing content is a glossary link.
   Write the helper (or use `${CLAUDE_SKILL_DIR}/scripts/glossary_links.py` if it exists) and make a
   marked copy of the script; run the converter on the **marked copy** in step 2:
   ```bash
   cat > /tmp/glossary_links.py <<'PY'
"""Green-highlighted text in a WTA script -> <wta-glossary-link>.
mark: copy the .docx, wrapping each run of green-highlighted text in ⟦ ⟧ markers.
link: in the converted .html, turn ⟦text⟧ into <wta-glossary-link term="…" text="…">
      (markers inside attribute values are just removed)."""
import re, sys, html
OPEN, CLOSE = "⟦", "⟧"
GREENS = {"green", "brightGreen"}

def mark(src, dst):
    import docx
    from docx.oxml.ns import qn
    d = docx.Document(src)
    def is_green(r):
        h = r.find(qn("w:rPr") + "/" + qn("w:highlight"))
        return h is not None and h.get(qn("w:val")) in GREENS
    def texts(r):
        return [t for t in r.iter(qn("w:t"))]
    n = 0
    for p in d.element.body.iter(qn("w:p")):
        runs = [r for r in p.iter(qn("w:r")) if texts(r)]
        i = 0
        while i < len(runs):
            if not is_green(runs[i]):
                i += 1; continue
            j = i
            while j + 1 < len(runs) and is_green(runs[j + 1]):
                j += 1
            group = runs[i:j + 1]
            if "".join(t.text or "" for r in group for t in texts(r)).strip():
                first, last = texts(group[0])[0], texts(group[-1])[-1]
                # keep leading/trailing spaces outside the link
                m = re.match(r"^(\s*)(.*)$", first.text or "", re.S)
                first.text = m.group(1) + OPEN + m.group(2)
                m = re.match(r"^(.*?)(\s*)$", last.text or "", re.S)
                last.text = m.group(1) + CLOSE + m.group(2)
                for t in (first, last):
                    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                n += 1
            i = j + 1
    d.save(dst)
    print(f"GLOSSARY marked {n} green-highlighted item(s)")

def term_of(text):
    plain = html.unescape(re.sub(r"<[^>]+>", "", text))
    return re.sub(r"[^0-9a-z]", "", plain.lower())

def glossary_terms(docx_path):
    """{term id: entry} from the script's 'Glossary term' table."""
    import docx
    d = docx.Document(docx_path)
    out = {}
    for t in d.tables:
        rows = t.rows
        if rows and rows[0].cells[0].text.strip().lower().startswith("glossary term"):
            for r in rows[1:]:
                k = r.cells[0].text.replace(OPEN, "").replace(CLOSE, "").strip()
                if k: out[term_of(k)] = k
    return out

def match_term(text, known):
    """Exact glossary id first; else the closest entry (plural/singular, leading 'the'); else derived."""
    t = term_of(text)
    if t in known: return t, "exact"
    t2 = re.sub(r"^the", "", t)
    if t2 in known: return t2, "fuzzy"
    best, score = None, 0
    for k in known:
        n = 0
        while n < min(len(t2), len(k)) and t2[n] == k[n]: n += 1
        if n >= 5 and n >= 0.8 * min(len(t2), len(k)) and (n > score or (n == score and len(k) < len(best))):
            best, score = k, n
    if best: return best, "fuzzy"
    return t, "none"

def link(path, docx_path=None):
    known = glossary_terms(docx_path) if docx_path else {}
    s = open(path, encoding="utf-8").read()
    # 1) strip markers inside attribute values (alt="", title="", src="" ...)
    s = re.sub(r'="([^"]*)"', lambda m: '="' + m.group(1).replace(OPEN, "").replace(CLOSE, "") + '"', s)
    # 2) convert the rest; inline tags inside the marked text are dropped from text=""
    linked, fuzzy, unmatched = [], [], []
    def rep(m):
        inner = m.group(1)
        text = html.unescape(re.sub(r"<[^>]+>", "", inner)).strip()
        if not text:
            return inner
        term, how = match_term(text, known) if known else (term_of(text), "none")
        linked.append(text)
        if how == "fuzzy": fuzzy.append(f"'{text}' -> {term} (glossary entry '{known[term]}')")
        if how == "none" and known: unmatched.append(f"'{text}' -> {term}")
        return f'<wta-glossary-link term="{term}" text="{html.escape(text, quote=True)}"></wta-glossary-link>'
    s = re.sub(OPEN + "(.*?)" + CLOSE, rep, s, flags=re.S)
    s = s.replace(OPEN, "").replace(CLOSE, "")
    open(path, "w", encoding="utf-8").write(s)
    print(f"GLOSSARY linked {len(linked)}: " + "; ".join(linked))
    for f in dict.fromkeys(fuzzy): print("CHECK glossary link matched to a similar glossary entry: " + f)
    for u in dict.fromkeys(unmatched): print("CHECK glossary link not in the script's glossary table, term id derived from the text: " + u)

if __name__ == "__main__":
    cmd = sys.argv[1]
    if cmd == "mark": mark(sys.argv[2], sys.argv[3])
    elif cmd == "link": link(sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else None)
PY
   python3 /tmp/glossary_links.py mark "<path/to/script.docx>" /tmp/<CODE>_marked.docx
   ```

1c. **If the user uploads an additional screenshot of the uploaded files (eg pictures, HTMLs), use it to update the source code with the correct file extension.** For example, if the script contains a reference to image N-AFRS-M1P01d, and the user uploads a screenshot of that image as N-AFRS-M1P01d.jpg, then run the converter with `--ext jpg` to update the source code to use the correct file extension. There are a couple of possible caveats to this that must be checked: first, the screenshot might not contain all the images in the script, so must not assume the extension for any images is the same as the ones in the screenshot. If there's an image in the script but not the screenshot, default to other rules in this process. Second, in the case of the HTML files, the markup code must use the file with _launcher.html at the end, as that is the launcher. If there's a HTML without a launcher, or there's a missing html from a screenshot with other HTMLs in it, flag with user and ask which one to use.

2. **Run the converter** (python-docx is preinstalled; if missing: `pip install python-docx --break-system-packages`):
   ```bash
   SCRIPT="${CLAUDE_SKILL_DIR}/scripts/docx_to_source.py"
   [ -f "$SCRIPT" ] || SCRIPT=$(find /sessions -name docx_to_source.py -path '*wta-docx-to-source*' 2>/dev/null | head -1)
   python3 "$SCRIPT" /tmp/<CODE>_marked.docx -o "<outputs folder>"
   ```
   Options: `--ext jpg` if the page's images are not .png; `--cdn <url>` to change the image host.
   The file is named `<CODE>.html` from the script's Code table (for example `FDF-SAS-A1-M3P03.html`).

2b. **Apply the case-study rule** (always run it; it is safe to run twice). A blue or grey callout
   (`alert-primary` or `highlight-box`) whose label line contains "Case study" becomes
   `<div class="case-study">`, with the title and first paragraph in one `<p>` joined by `&nbsp;<br>`:
   ```bash
   cat > /tmp/case_study_fix.py <<'PY'
import re, sys
p = sys.argv[1]; s = open(p).read()
def fix(m):
    inner = m.group(2)
    first = re.match(r'<p>(.*?)</p>\n?', inner, re.S)
    if not first or not re.search(r"case\s*stud(y|ies)", re.sub(r"<[^>]+>", "", first.group(1)), re.I):
        return m.group(0)
    title = first.group(1).strip()
    if not title.startswith("<strong>"): title = f"<strong>{title}</strong>"
    rest = inner[first.end():]
    nxt = re.match(r'<p>(.*?)</p>\n?', rest, re.S)
    if nxt:
        head = f"<p>{title}&nbsp;<br>{nxt.group(1)}</p>\n"; rest = rest[nxt.end():]
    else:
        head = f"<p>{title}</p>\n"
    return f'<div class="case-study">\n{head}{rest}</div>'
s = re.sub(r'<div class="(alert-primary|highlight-box)">\n(.*?)</div>', fix, s, flags=re.S)
open(p, "w").write(s)
PY
   python3 /tmp/case_study_fix.py "<outputs folder>/<CODE>.html"
   ```
   Result:
   ```html
   <div class="case-study">
   <p><strong>Case study: Vanguard Logistics</strong>&nbsp;<br>Vanguard Logistics, described by FreightWaves as …</p>
   </div>
   ```
   Red and green callouts are unchanged. Glossary links inside a case study come only from green highlights (step 2d).

2c. **Apply the SVG rule** (always run it; safe to run twice). Images inside a scroll animator or a
   text reveal are usually SVG uploads, so if no screenshot is available, their CDN path ends in `.svg`, not `.png`. Other images (banners,
   diagrams, standalone column sets, hotspots) keep `.png` (or the `--ext` value):
   ```bash
   cat > /tmp/widget_svg_fix.py <<'PY'
import re, sys
p = sys.argv[1]; s = open(p).read()
def to_svg(m):
    return re.sub(r'(https://cdn\.wisetechacademy\.com/[^"\s]+)\.(png|jpe?g)"', r'\1.svg"', m.group(0))
s = re.sub(r'<wta-scroll-animator\b.*?</wta-scroll-animator>', to_svg, s, flags=re.S)
s = re.sub(r'<wta-text-reveal\b.*?</wta-text-reveal>', to_svg, s, flags=re.S)
open(p, "w").write(s)
PY
   python3 /tmp/widget_svg_fix.py "<outputs folder>/<CODE>.html"
   ```
   If you rebuild a scroll animator or text reveal by hand, use `.svg` for its images too.

2d. **Turn the marks into glossary links** (pass the original .docx so term ids are matched to its glossary table):
   ```bash
   python3 /tmp/glossary_links.py link "<outputs folder>/<CODE>.html" "<path/to/script.docx>"
   ```
   Result:
   ```html
   <p>… the completion rules for the <wta-glossary-link term="customsdeclarationservicecds" text="Customs Declaration Service (CDS)"></wta-glossary-link>. …</p>
   ```
   - `text` is the highlighted text exactly as written; `term` is the glossary entry's id: lowercase,
     letters and digits only (for example "Special procedures" -> `specialprocedures`).
   - If the highlighted text matches a glossary-table entry, that entry's id is used. A close match
     (plural/singular, a leading "The") uses the entry's id and prints a CHECK line; no match derives
     the id from the text and prints a CHECK line. List both kinds in the reply.
   - Highlights inside alt text or other attributes are dropped (no link). Highlights in question
     tables, the glossary table, and instructions are not output anyway.

3. **Read the printed CHECK list** and the output file. Look at the output against the script once:
   every heading, paragraph, and widget in the body should appear in order. If a table was
   turned into a standard `<table>` and is actually a widget, fix it by hand using
   `references/widget-patterns.md`, then consider adding that widget to the converter.

4. **Deliver**: share the `.html` file(s) with present_files. In the reply, list the CHECK items
   as short bullets (hotspot positions, missing question IDs, missing links, and so on) so the LD
   knows what to finish in the editor. Keep it short.

## Question IDs

- The converter reads each question's ID from the **Question ID / Question id** row of its
  question table and writes it into `qids`. It never makes up an ID.
- A Quick Quiz question with a blank ID gets `qids="TODO"` and a CHECK item.
- **Updated script uploaded** (for example, the same code with IDs now filled in): run the
  converter again on the new .docx, overwrite the existing `<CODE>.html`, and compare it with the
  previous version (`diff`). Tell the LD which placeholders were replaced, for example:
  "`TODO` → `44117` (third Quick Quiz)". Also list any final-assessment IDs that are now filled in.

## What the converter does

| Script element | Output |
|---|---|
| Code / metadata table | Not output (gives the code and title) |
| Heading 2 / 3 / 4, Body Text, bullets, numbering | `h2` / `h3` / `h4`, `p`, `ul` / `ol` with bold, italic, links kept |
| Image / Graphic table | `<p><img src="https://cdn.wisetechacademy.com/resources/wta/{FOLDER}/{FILE}.png" class="block-center" alt="…"></p>` |
| Green-highlighted text (learner-facing) | `<wta-glossary-link term="{id}" text="{highlighted text}"></wta-glossary-link>` (steps 1b, 2d) |
| Call-out table or callout paragraph style | `<div class="alert-primary / alert-danger / alert-success">` |
| Blue or grey callout whose label line contains "Case study" | `<div class="case-study">` with `<p><strong>Case study: …</strong>&nbsp;<br>first paragraph…</p>` |
| Scroll animator | one `wta-scroll-animator` + `wta-column-set` per point, icon in the 25% column, image path ends in `.svg` |
| Image hotspot + Marker tables | `wta-hotspot` with one `wta-hotspot-marker` per marker (x/y = 50/50 placeholders) |
| Text reveal + Card tables | `wta-text-reveal` with `wta-text-reveal-card`s, card image path ends in `.svg` |
| "DIY: Dropdown (this and others)" comment | that heading and the following section headings become `wta-expander`s |
| HTML import | `wta-html-import` with the URL from the script |
| Quick Quiz question tables | `wta-question qqtype="QuickQuiz" qids="…"` at the table position, or at "Add Quick Quiz N here"; back-to-back IDs are combined |
| List block (Course resource) | `<ul class="course-resource">` |
| Final-assessment questions, glossary table, image briefs, instructions ("Please write…", "Add … here") | Not output |
| Any other table | `<table class="default">` and a CHECK item |

Both script template versions are supported: the older one ("Graphic" tables, blue Code
value cell, "Question use: Quick Quiz 1" with an "Add Quick Quiz 1 here" line) and the newer
one ("Image" tables with CCA dimensions, "Question use: In quick quiz").

## Rules the team agreed

- **Images**: full CDN URL `https://cdn.wisetechacademy.com/resources/wta/{FOLDER}/{FILE}.{ext}`,
  where FOLDER is the file name up to its `-M#P#` part. Reused assets from another course keep their own folder.
- **Widget images are SVG**: images in a scroll animator or text reveal use `.svg`; all other images use `.png` (or `--ext`).
- **Glossary**: only green-highlighted learner-facing text becomes `<wta-glossary-link term="…" text="…">` (steps 1b and 2d). Do not add glossary links anywhere else.
- **Quick Quiz**: placed where the table (or "Add Quick Quiz N here" line) sits in the script.
- **Output**: the HTML source file only.
- Never invent content, IDs, URLs, or coordinates. Use a `TODO` placeholder and list it as a CHECK item.

## Updating the mappings

Class mappings (callout types/styles), skipped instruction lines, and stop headings are
constants at the top of `scripts/docx_to_source.py`. When the script template adds a new
widget table, add a branch in `Converter.handle_table` keyed on the table's first-cell text,
and add its pattern to `references/widget-patterns.md`.
