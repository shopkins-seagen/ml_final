import pathlib

base = pathlib.Path(r'c:/dev/stevens/data_analytics/project/ML_domain_classifier-')

# Write pfizer.scss
scss = (
    "/*-- scss:defaults --*/\n\n"
    "$body-bg: #FFFFFF;\n"
    "$body-color: #1A1A2E;\n"
    "$link-color: #0053A0;\n"
    "$selection-bg: #CCE0F5;\n"
    "$presentation-heading-color: #0053A0;\n"
    '$presentation-heading-font: "Segoe UI", Arial, sans-serif;\n'
    '$font-family-sans-serif: "Segoe UI", Arial, sans-serif;\n'
    "$code-block-bg: #EEF3FA;\n\n"
    "/*-- scss:rules --*/\n\n"
    ".reveal,\n.reveal .slides,\n.reveal .slides section {\n"
    "  background-color: #FFFFFF !important;\n"
    "  color: #1A1A2E !important;\n}\n\n"
    ".reveal .slides section {\n"
    "  padding: 20px 48px !important;\n"
    "  box-sizing: border-box;\n}\n\n"
    ".reveal h1,\n.reveal h2 {\n"
    "  color: #0053A0;\n"
    "  border-bottom: 3px solid #00A3E0;\n"
    "  padding-bottom: 8px;\n"
    "  margin-bottom: 16px;\n}\n\n"
    ".reveal h3,\n.reveal h4 {\n  color: #001F5B;\n}\n\n"
    ".reveal p,\n.reveal li,\n.reveal span {\n  color: #1A1A2E !important;\n}\n\n"
    ".reveal .title-slide,\n.reveal .title-slide h1,\n"
    ".reveal .title-slide h2,\n.reveal .title-slide p {\n"
    "  background-color: #FFFFFF !important;\n"
    "  color: #0053A0 !important;\n}\n\n"
    ".reveal table {\n  width: 100%;\n  border-collapse: collapse;\n  margin: 12px 0;\n}\n\n"
    ".reveal table thead tr {\n  background-color: #0053A0;\n}\n\n"
    ".reveal table thead tr th {\n  color: #FFFFFF !important;\n  padding: 8px 12px;\n}\n\n"
    ".reveal table tbody tr:nth-child(even) {\n  background-color: #E8F1FA;\n}\n\n"
    ".reveal table tbody tr td {\n  color: #1A1A2E !important;\n  padding: 6px 12px;\n}\n\n"
    ".reveal pre {\n"
    "  border-left: 4px solid #0053A0;\n"
    "  background-color: #EEF3FA !important;\n"
    "  color: #1A1A2E !important;\n"
    "  padding: 12px;\n"
    "  border-radius: 4px;\n}\n\n"
    ".reveal code {\n  background-color: #EEF3FA !important;\n  color: #1A1A2E !important;\n}\n\n"
    ".reveal .progress {\n  color: #0053A0;\n  height: 4px;\n}\n\n"
    ".reveal .slide-number {\n"
    "  background-color: #0053A0;\n"
    "  color: #FFFFFF !important;\n"
    "  border-radius: 4px;\n"
    "  padding: 2px 6px;\n}\n\n"
    ".reveal .footer {\n"
    "  color: #0053A0 !important;\n"
    "  font-size: 0.7em;\n"
    "  border-top: 2px solid #00A3E0;\n"
    "  padding-top: 4px;\n}\n\n"
    ".callout {\n"
    "  border-left-color: #0053A0 !important;\n"
    "  background-color: #EEF3FA !important;\n"
    "  color: #1A1A2E !important;\n}\n\n"
    ".reveal ul,\n.reveal ol {\n  margin-left: 1.2em;\n  line-height: 1.7;\n}\n"
)

(base / 'pfizer.scss').write_text(scss, encoding='utf-8')
print('pfizer.scss written:', (base / 'pfizer.scss').stat().st_size, 'bytes')

# Fix encoding artifacts in slides.qmd
slides_path = base / 'slides.qmd'
content = slides_path.read_text(encoding='utf-8', errors='replace')

for bad, good in [
    ('\u00e2\u0080\u0093', '-'),
    ('\u00e2\u0080\u0094', '-'),
    ('\u00e2\u0080\u0099', "'"),
    ('\u00e2\u0080\u009c', '"'),
    ('\u00e2\u0080\u009d', '"'),
    ('\u00c2\u00b7', '\u00b7'),
    ('â€"', '-'),
    ("â€™", "'"),
    ('â€œ', '"'),
    ('â€', '"'),
    ('Â·', '·'),
    ('Â', ''),
]:
    content = content.replace(bad, good)

slides_path.write_text(content, encoding='utf-8')
print('slides.qmd fixed:', slides_path.stat().st_size, 'bytes')
print('Done!')
