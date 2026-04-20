import re

special_slides = [
    "Project Objective",
    "KNN Interpretation",
    "CART Interpretation",
    "Model Comparison"
]

def extract_sections(content):
    # Returns a list of (heading, section_text)
    pattern = r"(## .+?)(?=\n## |\Z)"
    matches = re.findall(pattern, content, re.DOTALL)
    sections = []
    for match in matches:
        heading = re.match(r"## (.+)", match)
        if heading:
            sections.append((heading.group(1).strip(), match))
    return sections

def extract_smaller_table(section):
    table_pattern = r"::: \{\.smaller\}(.*?):::"  # non-greedy
    match = re.search(table_pattern, section, re.DOTALL)
    return match.group(1) if match else None

def replace_section(slides_content, heading, new_section, special_table=False):
    pattern = rf"(## {re.escape(heading)}\n.*?)(?=\n## |\Z)"
    if special_table:
        table_pattern = r"(## " + re.escape(heading) + r"\n.*?::: \{\.smaller\})(.*?)(:::)"
        new_table = extract_smaller_table(new_section)
        if new_table:
            slides_content = re.sub(
                table_pattern,
                rf"\1{new_table}\3",
                slides_content,
                flags=re.DOTALL
            )
    else:
        slides_content = re.sub(pattern, new_section, slides_content, flags=re.DOTALL)
    return slides_content

with open("index.qmd", encoding="utf-8") as f:
    index_content = f.read()
with open("slides2.qmd", encoding="utf-8") as f:
    slides_content = f.read()

# Get all slide headings in slides2.qmd
slide_headings = set(re.findall(r"^## (.+)$", slides_content, re.MULTILINE))

# Get all sections from index.qmd
index_sections = extract_sections(index_content)

for heading, new_section in index_sections:
    if heading in slide_headings:
        slides_content = replace_section(
            slides_content,
            heading,
            new_section,
            special_table=(heading in special_slides)
        )
    else:
        # Heading not present, so append section to the end
        slides_content += f"\n\n{new_section.strip()}\n"

with open("slides2.qmd", "w", encoding="utf-8", newline="\n") as f:
    f.write(slides_content)

print("slides2.qmd updated: all sections from index.qmd are now included.")