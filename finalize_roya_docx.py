#!/usr/bin/env python3
"""
Finalize the deliverable Word document: gasoline_demand_paper(roya).docx.

Base: roya_original_backup.docx — an unaltered copy of the author's own
document as saved by Word (their byline, their course line, their text).

This script changes exactly what was asked:

  1. GitHub link: every instance of github.com/[username]/gasoline-demand
     (displayed text AND hyperlink targets, including the %5b/%5d
     percent-encoded form Word writes) becomes
     github.com/RoyaRezaie/gasoline-demand.
  2. Footnotes: the seven source mentions (SSRN survey paper, EIA, FRED,
     Wooldridge, Hughes et al., Havranek et al., Espey) become regular
     Word footnotes at the bottom of the page, each carrying the full
     bibliographic citation, anchored at the first mention in the main
     text.

Nothing else is touched.

Run:  python3 paper/finalize_roya_docx.py
Output: paper/gasoline_demand_paper(roya).docx
"""

import copy
import os
import re
import zipfile

from lxml import etree

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "roya_original_backup.docx")
OUT = os.path.join(HERE, "gasoline_demand_paper(roya).docx")

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_R = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"

OLD_URL = "github.com/[username]/gasoline-demand"
NEW_URL = "github.com/RoyaRezaie/gasoline-demand"
OLD_URL_ENC = "github.com/%5busername%5d/gasoline-demand"

# ── the seven source footnotes, in document order ─────────────────────
# (paragraph locator, anchor text or None = end of paragraph, citation)
FOOTNOTES = [
    ("business question addressed in this paper", None,
     "Using Supply and Demand Functions: A Survey by Data Availability. "
     "SSRN Electronic Journal. "
     "https://papers.ssrn.com/sol3/papers.cfm?abstract_id=7399858 "
     "(starting point for this project)."),
    ("Administration (EIA)", "Administration (EIA)",
     "U.S. Energy Information Administration. Petroleum and Other Liquids: "
     "Supply and Disposition, and Retail Prices (monthly series, downloaded "
     "October 2026). https://www.eia.gov/dnav/pet/"),
    ("Federal Reserve Bank of St. Louis, I estimate",
     "Federal Reserve Bank of St. Louis",
     "Federal Reserve Bank of St. Louis. FRED: CPIAUCSL, DSPIC96, POPTHM, "
     "UNRATE, DCOILWTICO (downloaded October 2026). "
     "https://fred.stlouisfed.org/"),
    ("adds the lagged level of ln Q.", "adds the lagged level of ln Q.",
     "Wooldridge, J. M. (2020). Introductory Econometrics: A Modern "
     "Approach (7th ed.). Cengage Learning."),
    ("Sperling (2008)” report", "Sperling (2008)”",
     "Hughes, J. E., Knittel, C. R., & Sperling, D. (2008). Evidence of a "
     "shift in the short-run price elasticity of gasoline demand. "
     "The Energy Journal, 29(1), 113–134."),
    ("Janda (2012), correcting", "Janda (2012)",
     "Havranek, T., Irsova, Z., & Janda, K. (2012). Demand for gasoline is "
     "more price-inelastic than commonly thought. Energy Economics, "
     "34(1), 201–207."),
    ("Espey (1998) is the classic", "Espey (1998)",
     "Espey, M. (1998). Gasoline demand revisited: An international "
     "meta-analysis of cross-country elasticity estimates. "
     "Energy Economics, 20(3), 273–295."),
]

FOOTNOTE_STYLES = f'''<w:style w:type="paragraph" w:styleId="FootnoteText">
 <w:name w:val="footnote text"/><w:basedOn w:val="Normal"/>
 <w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/>
 <w:pPr><w:spacing w:before="0" w:after="0" w:line="240"
   w:lineRule="auto"/></w:pPr>
 <w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"
   w:cs="Times New Roman"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>
</w:style>
<w:style w:type="character" w:styleId="FootnoteReference">
 <w:name w:val="footnote reference"/><w:uiPriority w:val="99"/>
 <w:semiHidden/><w:unhideWhenUsed/>
 <w:rPr><w:vertAlign w:val="superscript"/></w:rPr>
</w:style>'''


def q(tag):
    return f"{{{W}}}{tag}"


def para_text(p):
    return "".join(t.text or "" for t in p.iter(q("t")))


def run_text(r):
    return "".join(t.text or "" for t in r.iter(q("t")))


def normalize_run(r):
    """Guarantee a single w:t carrying the run's full text."""
    ts = r.findall(q("t"))
    if len(ts) > 1:
        full = "".join(t.text or "" for t in ts)
        for t in ts[1:]:
            r.remove(t)
        ts[0].text = full
        ts[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    elif len(ts) == 1 and ts[0].text is None:
        ts[0].text = ""
    return ts[0]


def replace_across(ts, old, new):
    """Replace `old` with `new` across the concatenated text of a list of
    w:t nodes (handles Word splitting a URL over several runs).
    Returns the number of replacements."""
    if not ts:
        return 0
    texts = [t.text or "" for t in ts]
    full = "".join(texts)
    starts, i = [], 0
    while True:
        i = full.find(old, i)
        if i < 0:
            break
        starts.append(i)
        i += len(old)
    for a in reversed(starts):
        b = a + len(old)
        bounds, p = [], 0
        for tx in texts:
            bounds.append((p, p + len(tx)))
            p += len(tx)
        first = next(k for k, (s, e) in enumerate(bounds) if s <= a < e)
        last = next(k for k, (s, e) in enumerate(bounds) if s < b <= e)
        pre = texts[first][:a - bounds[first][0]]
        suf = texts[last][b - bounds[last][0]:]
        if first == last:
            texts[first] = pre + new + suf
        else:
            texts[first] = pre + new
            for k in range(first + 1, last):
                texts[k] = ""
            texts[last] = suf
    for t, tx in zip(ts, texts):
        t.text = tx
        if tx[:1].isspace() or tx[-1:].isspace():
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return len(starts)


def make_ref_run(note_id):
    xml = (f'<w:r xmlns:w="{W}"><w:rPr>'
           f'<w:rStyle w:val="FootnoteReference"/></w:rPr>'
           f'<w:footnoteReference w:id="{note_id}"/></w:r>')
    return etree.fromstring(xml.encode())


def insert_ref(p, anchor, note_id):
    """Insert a footnote-reference run after `anchor`
    (at paragraph end when anchor is None), splitting a run if needed."""
    runs = [r for r in p.iter(q("r"))
            if r.find(q("footnoteReference")) is None]
    if anchor is None:
        runs[-1].addnext(make_ref_run(note_id))
        return
    spans, pos = [], 0
    for r in runs:
        t = run_text(r)
        spans.append((r, t, pos, pos + len(t)))
        pos += len(t)
    full = "".join(t for _, t, _, _ in spans)
    i = full.find(anchor)
    if i < 0:
        raise SystemExit(f"anchor not found in paragraph: {anchor!r}")
    cut = i + len(anchor)              # insert AFTER the anchor
    for r, t, s, e in spans:
        if s < cut <= e:
            k = cut - s                # offset inside this run
            if k == len(t):
                r.addnext(make_ref_run(note_id))
            else:
                r2 = copy.deepcopy(r)
                normalize_run(r).text = t[:k]
                normalize_run(r2).text = t[k:]
                r.addnext(r2)
                r.addnext(make_ref_run(note_id))
            return
    raise SystemExit(f"could not place reference for {anchor!r}")


def main():
    zin = zipfile.ZipFile(SRC)
    parts = {n: zin.read(n) for n in zin.namelist()}
    infos = {i.filename: i for i in zin.infolist()}
    zin.close()

    # ── 1. GitHub link: displayed text (incl. run-split occurrence) ──
    doc = etree.fromstring(parts["word/document.xml"])
    text_hits = 0
    for h in doc.iter(q("hyperlink")):
        text_hits += replace_across(list(h.iter(q("t"))), OLD_URL, NEW_URL)
    for p in doc.iter(q("p")):
        ts = [t for t in p.iter(q("t"))
              if not any(a.tag == q("hyperlink")
                         for a in t.iterancestors())]
        text_hits += replace_across(ts, OLD_URL, NEW_URL)
    assert text_hits == 4, f"displayed URL instances: {text_hits} (expected 4)"

    # ── 2. GitHub link: hyperlink targets ────────────────────────────
    rels = etree.fromstring(parts["word/_rels/document.xml.rels"])
    link_hits = 0
    for rel in rels.iter(f"{{{PKG_R}}}Relationship"):
        tgt = rel.get("Target", "")
        if "github.com" in tgt:
            link_hits += 1
            rel.set("Target", tgt.replace(OLD_URL_ENC, NEW_URL)
                    .replace(OLD_URL, NEW_URL))
    assert link_hits == 2, f"hyperlink targets: {link_hits} (expected 2)"
    assert not any(OLD_URL_ENC in (r.get("Target") or "")
                   for r in rels.iter(f"{{{PKG_R}}}Relationship"))
    parts["word/_rels/document.xml.rels"] = etree.tostring(
        rels, xml_declaration=True, encoding="UTF-8", standalone=True)

    # ── 3. footnote references in the body ───────────────────────────
    paragraphs = list(doc.iter(q("p")))
    for n, (locator, anchor, _cite) in enumerate(FOOTNOTES, start=1):
        target = next((p for p in paragraphs if locator in para_text(p)),
                      None)
        if target is None:
            raise SystemExit(f"paragraph not found: {locator!r}")
        insert_ref(target, anchor, n)
    order = [int(r.get(q("id"))) for r in doc.iter(q("footnoteReference"))]
    assert order == list(range(1, len(FOOTNOTES) + 1)), f"ref order: {order}"
    parts["word/document.xml"] = etree.tostring(
        doc, xml_declaration=True, encoding="UTF-8", standalone=True)

    # ── 4. footnote texts (Word's boilerplate part already exists) ───
    fns = etree.fromstring(parts["word/footnotes.xml"])
    existing = [int(f.get(q("id"))) for f in fns.iter(q("footnote"))]
    for n, (_loc, _a, cite) in enumerate(FOOTNOTES, start=1):
        assert n not in existing, f"footnote {n} already present"
        xml = (
            f'<w:footnote xmlns:w="{W}" w:id="{n}"><w:p><w:pPr>'
            f'<w:pStyle w:val="FootnoteText"/></w:pPr><w:r><w:rPr>'
            f'<w:rStyle w:val="FootnoteReference"/></w:rPr>'
            f'<w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> '
            f'{cite.replace("&", "&amp;").replace("<", "&lt;")}'
            f'</w:t></w:r></w:p></w:footnote>')
        fns.append(etree.fromstring(xml.encode()))
    parts["word/footnotes.xml"] = etree.tostring(
        fns, xml_declaration=True, encoding="UTF-8", standalone=True)

    # ── 5. content type + relationship: already wired by Word ────────
    ct = parts["[Content_Types].xml"].decode("utf-8")
    assert "footnotes+xml" in ct, "content type for footnotes missing"
    assert any(r.get("Type", "").endswith("/footnotes")
               for r in rels.iter(f"{{{PKG_R}}}Relationship")), \
        "document -> footnotes relationship missing"

    # ── 6. footnote styles (absent from the author's copy) ───────────
    styles = etree.fromstring(parts["word/styles.xml"])
    have = {s.get(q("styleId")) for s in styles.iter(q("style"))}
    if "FootnoteText" not in have or "FootnoteReference" not in have:
        frag = etree.fromstring(
            f'<w:root xmlns:w="{W}"><w:body>{FOOTNOTE_STYLES}</w:body>'
            f'</w:root>')
        for child in list(frag.find(q("body"))):
            styles.append(child)
    parts["word/styles.xml"] = etree.tostring(
        styles, xml_declaration=True, encoding="UTF-8", standalone=True)

    # ── 7. write the deliverable ─────────────────────────────────────
    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in parts.items():
            zout.writestr(infos[name], data)
    print(f"Wrote {OUT}")
    print(f"  github link: {text_hits} displayed + {link_hits} hyperlink "
          f"targets -> {NEW_URL}")
    print(f"  footnotes: {len(FOOTNOTES)} "
          "(SSRN, EIA, FRED, Wooldridge, Hughes, Havranek, Espey)")


if __name__ == "__main__":
    main()
