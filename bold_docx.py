#!/usr/bin/env python3

import sys
import zipfile
from lxml import etree as ET
from io import BytesIO
import os

def split_on_bold_markers(text):
    return text.split("**")

def create_rich_text_children_docx(r_element, original_text, namespace):
    parts = split_on_bold_markers(original_text)

    if len(parts) % 2 != 1:
        print("Unmatched ** marker found. Skipping paragraph.")
        return None

    new_children = []

    original_rpr_children = r_element.findall(f"./rPr/*")

    for i in range(len(parts)):
        new_child = ET.Element(f"{{{namespace}}}r")
        # new_child.set("rsidRPr", r_element.get("rsidRPr"))

        if len(parts[i]) == 0:
            continue

        if i % 2 == 1: # Should be bold
            rPr_child = ET.SubElement(new_child, f"{{{namespace}}}rPr")

            for original_rpr_child in original_rpr_children:
                rPr_child.append(original_rpr_child)

            bold_child = ET.SubElement(rPr_child, f"{{{namespace}}}b")
            bCs_child = ET.SubElement(rPr_child, f"{{{namespace}}}bCs")

        text_child = ET.SubElement(new_child, f"{{{namespace}}}t")
        text_child.text = parts[i]

        new_children.append(new_child)

    return new_children


def get_namespace(root):
    return root.tag.split("}")[0].strip("{") if "}" in root.tag else ""


# Applies bold to the paragraphs in document.xml
def bold_document_xml(file_dict):
    target_sheet = "word/document.xml"

    file_data = file_dict.get(target_sheet)
    if file_data is None:
        return

    root = ET.fromstring(file_data)

    namespace = get_namespace(root)

    for paragraph in root.iter(f"{{{namespace}}}p"):
        r_elements = paragraph.findall(f"{{{namespace}}}r")

        if len(r_elements) == 0:
            continue

        original_text = ""

        for r_element in r_elements:
            text_child = r_element.find(f"{{{namespace}}}t")

            if text_child is None:
                continue

            original_text += text_child.text

        if "**" not in original_text:
            continue

        new_children = create_rich_text_children_docx(r_elements[0], original_text, namespace)
        if new_children is None:
            continue

        original_attributes = paragraph.items()

        paragraph.clear()
        for key, val in original_attributes:
            paragraph.set(key, val)

        for child in new_children:
            paragraph.append(child)

        print(f"Added bold to paragraph {original_text}")

    modified_xml = BytesIO()
    tree = ET.ElementTree(root)
    tree.write(modified_xml, encoding="utf-8", xml_declaration=True)
    file_dict[target_sheet] = modified_xml.getvalue()


def main():
    argc = len(sys.argv)

    if argc < 2:
        print("Invalid number of arguments, expected 1: .docx filename")
        sys.exit(1)

    input_file_path = sys.argv[1]
    if not os.path.isfile(input_file_path) or os.path.splitext(input_file_path)[1] != ".docx":
        print(f"Error: {input_file_path} is not a valid .docx file.")
        sys.exit(1)
    

    directory = os.path.dirname(input_file_path)

    output_file_name = os.path.basename(input_file_path)
    output_file_path = os.path.join(directory, "BOLD_" + output_file_name)

    file_dict = {}
    with zipfile.ZipFile(input_file_path, "r") as zip_ref:
        file_dict = {name: zip_ref.read(name) for name in zip_ref.namelist()}


    bold_document_xml(file_dict)

    with zipfile.ZipFile(output_file_path, 'w') as zip_out:
        for file_name, data in file_dict.items():
            zip_out.writestr(file_name, data)

    print(f"Saved updated .docx to {output_file_path}")


if __name__ == "__main__":
    main()
