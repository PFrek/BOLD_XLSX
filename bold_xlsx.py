#!/usr/bin/env python3

import sys
import zipfile
from lxml import etree as ET
from io import BytesIO
import os

MAX_SHEET_NUM = 100

def split_on_bold_markers(text):
    return text.split("**")

def create_rich_text_children(text, namespace):
    parts = split_on_bold_markers(text)

    if len(parts) % 2 != 1:
        print("Unmatched ** marker found. Skipping cell.")
        return None

    new_children = []

    for i in range(len(parts)):
        new_child = ET.Element(f"{{{namespace}}}r")

        if len(parts[i]) == 0:
            continue

        if i % 2 == 1: # Should be bold
            rPr_child = ET.SubElement(new_child, f"{{{namespace}}}rPr")
            bold_child = ET.SubElement(rPr_child, f"{{{namespace}}}b")

        text_child = ET.SubElement(new_child, f"{{{namespace}}}t")
        text_child.text = parts[i]

        new_children.append(new_child)

    return new_children


def get_namespace(root):
    return root.tag.split("}")[0].strip("{") if "}" in root.tag else ""


# Applies bold to the strings in sharedStrings.xml
def bold_shared_strings_xml(file_dict):
    target_sheet = "xl/sharedStrings.xml"

    root = ET.fromstring(file_dict[target_sheet])

    namespace = get_namespace(root)

    for cell in root.iter(f"{{{namespace}}}si"):
        text_child = cell.find(f"{{{namespace}}}t")

        if text_child is None:
            print("No <t> element found in cell. Skipping.")
            continue

        original_text = text_child.text

        if "**" not in text_child.text:
            continue

        new_children = create_rich_text_children(text_child.text, namespace)
        if new_children is None:
            continue

        cell.clear()
        for child in new_children:
            cell.append(child)

        print(f"Added bold to shared string {original_text}")

    modified_xml = BytesIO()
    tree = ET.ElementTree(root)
    tree.write(modified_xml, encoding="utf-8", xml_declaration=True)
    file_dict[target_sheet] = modified_xml.getvalue()


# Applies bold to string cells in each worksheet
# This was needed because cells that contain formulas did not use the sharedStrings.xml
# definitions.
def bold_references_xml(file_dict):
    formula_cells_for_removal = []

    for i in range(1, MAX_SHEET_NUM):
        target_sheet = f"xl/worksheets/sheet{i}.xml"

        file_data = file_dict.get(target_sheet)
        if file_data is None:
            break # No more sheets to read

        root = ET.fromstring(file_data)

        namespace = get_namespace(root)

        for cell in root.iter(f"{{{namespace}}}c"):
            if cell.get("t") != "str":
                continue

            v_child = cell.find(f"{{{namespace}}}v")

            if v_child is None:
                print("No <v> element found in cell. Skipping.")
                continue

            original_text = v_child.text

            if "**" not in v_child.text:
                continue


            # If cell contained formula reference, save its number for removal from calcChain.xml
            f_element = cell.find(f"{{{namespace}}}f")
            if f_element is not None:
                formula_cells_for_removal.append((i, cell.get("r")))

            new_children = create_rich_text_children(v_child.text, namespace)

            original_attributes = cell.items()

            cell.clear()

            # Reapply old attributes
            for key, val in original_attributes:
                cell.set(key, val)

            cell.set("t", "inlineStr")

            si_child = ET.Element(f"{{{namespace}}}is")
            for child in new_children:
                si_child.append(child)

            cell.append(si_child)

            print(f"Added bold to original text {original_text}")

        modified_xml = BytesIO()
        tree = ET.ElementTree(root)
        tree.write(modified_xml, encoding="utf-8", xml_declaration=True)
        file_dict[target_sheet] = modified_xml.getvalue()

    remove_formula_references_xml(file_dict, formula_cells_for_removal)


# This removes the references to formulas that were changed by bold_references_xml
# Since that function replaced the references in sheet#.xml for inlineStrings, then
# the formulas also need to be removed from the calcChain.xml file
def remove_formula_references_xml(file_dict, cells):
    target_sheet = f"xl/calcChain.xml"

    file_data = file_dict.get(target_sheet)
    if file_data == None:
        return # No calcChain file

    root = ET.fromstring(file_data)

    namespace = get_namespace(root)

    for removal_info in cells:
        cell_sheet_i = removal_info[0]
        cell_row = removal_info[1]

        xpath = f".//{{{namespace}}}c[@r='{cell_row}']"

        c_elements = root.findall(xpath)
        element = None

        for c in c_elements:
            if c.get('i') == str(cell_sheet_i):
                element = c
                break

        if element is not None:
            root.remove(element)
            print(f"Removed formula reference to {cell_sheet_i}:{cell_row} from calcChain.xml")


    n_children = len(root.findall("*"))

    if n_children == 0:
        del file_dict[target_sheet]
        print("No more formula references, deleting calcChain.xml from archive")
    else:
        modified_xml = BytesIO()
        tree = ET.ElementTree(root)
        tree.write(modified_xml, encoding="utf-8", xml_declaration=True)
        file_dict[target_sheet] = modified_xml.getvalue()





def main():
    argc = len(sys.argv)

    if argc < 2:
        print("Invalid number of arguments, expected 1: .xlsx filename")
        sys.exit(1)

    input_file_path = sys.argv[1]
    if not os.path.isfile(input_file_path) or os.path.splitext(input_file_path)[1] != ".xlsx":
        print(f"Error: {input_file_path} is not a valid .xlsx file.")
        sys.exit(1)
    

    directory = os.path.dirname(input_file_path)

    output_file_name = os.path.basename(input_file_path)
    output_file_path = os.path.join(directory, "BOLD_" + output_file_name)

    file_dict = {}
    with zipfile.ZipFile(input_file_path, "r") as zip_ref:
        file_dict = {name: zip_ref.read(name) for name in zip_ref.namelist()}


    bold_shared_strings_xml(file_dict)

    bold_references_xml(file_dict)

    with zipfile.ZipFile(output_file_path, 'w') as zip_out:
        for file_name, data in file_dict.items():
            zip_out.writestr(file_name, data)

    print(f"Saved updated .xlsx to {output_file_path}")


if __name__ == "__main__":
    main()
