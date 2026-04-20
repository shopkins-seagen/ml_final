import re

# Slides that use the .smaller table format
special_slides = [
    "Project Objective",
    "KNN Interpretation",
    "CART Interpretation",
    "Model Comparison"
]

def extract_section(content, heading):
    # Matches a section starting with '## heading' up to the next '## ' or end of file
    pattern = rf"(## {re.escape(heading)}\n.*?)(?=\n## |\Z)"
    match = re.search(pattern, content, re.DOTALL)
    return match.group(1) if match else None

def extract_smaller_table(section):
    # Extracts the content inside the .smaller table
    table_pattern = r"::: \{\.smaller\}(.*?):::"  # non-greedy
    match = re.search(table_pattern, section, re.DOTALL)
    return match.group(1) if match else None

def replace_section(slides_content, heading, new_section, special_table=False):
    # Replace the section in slides_content with new_section
    pattern = rf"(## {re.escape(heading)}\n.*?)(?=\n## |\Z)"
    if special_table:
        # Only replace the content inside the .smaller table
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

# List all slide headings in slides2.qmd
slide_headings = re.findall(r"^## (.+)$", slides_content, re.MULTILINE)

for heading in slide_headings:
    new_section = extract_section(index_content, heading)
    if new_section:
        slides_content = replace_section(
            slides_content,
            heading,
            new_section,
            special_table=(heading in special_slides)
        )

with open("slides2.qmd", "w", encoding="utf-8", newline="\n") as f:
    f.write(slides_content)

print("slides2.qmd updated with all code/output from index.qmd.")