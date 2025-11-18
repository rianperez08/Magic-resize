import sys
import shutil
import zipfile
import os
import xml.etree.ElementTree as ET
from copy import deepcopy

# Namespaces used in PPTX Open XML
NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

ET.register_namespace("p", NS["p"])
ET.register_namespace("r", NS["r"])


def extract_pptx(pptx_path, extract_dir):
    with zipfile.ZipFile(pptx_path, "r") as z:
        z.extractall(extract_dir)


def rezip_pptx(folder, out_path):
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        for root, dirs, files in os.walk(folder):
            for f in files:
                full = os.path.join(root, f)
                rel = os.path.relpath(full, folder)
                z.write(full, rel)


def get_slide_order(presentation_xml_path):
    tree = ET.parse(presentation_xml_path)
    root = tree.getroot()
    sldIdLst = root.find("p:sldIdLst", NS)
    if sldIdLst is None:
        return [], tree, root, sldIdLst

    slide_ids = []
    for sldId in sldIdLst.findall("p:sldId", NS):
        rId = sldId.attrib.get(f"{{{NS['r']}}}id")
        slide_ids.append((sldId, rId))
    return slide_ids, tree, root, sldIdLst


def load_rels(rels_path):
    if not os.path.exists(rels_path):
        return None, None, None
    tree = ET.parse(rels_path)
    root = tree.getroot()
    return tree, root, root.findall("r:Relationship", NS)


def next_id(existing_ids):
    """Simple helper to get next numeric ID as string."""
    if not existing_ids:
        return "1"
    return str(max(int(x) for x in existing_ids) + 1)


def add_morph_transition_to_slide(slide_xml_path):
    tree = ET.parse(slide_xml_path)
    root = tree.getroot()

    # p:sld root
    transition = root.find("p:transition", NS)
    if transition is None:
        transition = ET.SubElement(root, f"{{{NS['p']}}}transition")
    # clear existing children
    for child in list(transition):
        transition.remove(child)
    # add morph child
    ET.SubElement(transition, f"{{{NS['p']}}}morph")

    tree.write(slide_xml_path, xml_declaration=True, encoding="utf-8")


def main(original_pptx, resized_pptx, output_pptx):
    work_dir = os.path.join(
        os.path.dirname(output_pptx),
        "_morph_work",
    )
    if os.path.exists(work_dir):
        shutil.rmtree(work_dir)
    os.makedirs(work_dir, exist_ok=True)

    # Unzip original and resized
    orig_dir = os.path.join(work_dir, "orig")
    res_dir = os.path.join(work_dir, "resized")
    out_dir = os.path.join(work_dir, "out")

    extract_pptx(original_pptx, orig_dir)
    extract_pptx(original_pptx, out_dir)  # base from original
    extract_pptx(resized_pptx, res_dir)

    # Load slide orders
    orig_pres_xml = os.path.join(orig_dir, "ppt", "presentation.xml")
    res_pres_xml = os.path.join(res_dir, "ppt", "presentation.xml")
    out_pres_xml = os.path.join(out_dir, "ppt", "presentation.xml")

    orig_slds, _, _, _ = get_slide_order(orig_pres_xml)
    res_slds, _, _, _ = get_slide_order(res_pres_xml)
    out_slds, out_tree, out_root, out_sldIdLst = get_slide_order(out_pres_xml)

    # Load presentation relationships for 'out'
    out_rels_path = os.path.join(out_dir, "ppt", "_rels", "presentation.xml.rels")
    out_rels_tree, out_rels_root, out_rels = load_rels(out_rels_path)

    # Load rel maps for original and resized
    orig_rels_path = os.path.join(orig_dir, "ppt", "_rels", "presentation.xml.rels")
    _, orig_rels_root, orig_rels = load_rels(orig_rels_path)

    res_rels_path = os.path.join(res_dir, "ppt", "_rels", "presentation.xml.rels")
    _, res_rels_root, res_rels = load_rels(res_rels_path)

    def relmap(rel_root):
        m = {}
        if rel_root is None:
            return m
        for rel in rel_root.findall("r:Relationship", NS):
            m[rel.attrib["Id"]] = rel.attrib["Target"]
        return m

    orig_rel_map = relmap(orig_rels_root)
    res_rel_map = relmap(res_rels_root)

    # Existing rel IDs
    existing_ids = []
    if out_rels is not None:
        for rel in out_rels:
          rid = rel.attrib.get("Id", "")
          if rid.startswith("rId"):
            try:
              existing_ids.append(rid[3:])
            except ValueError:
              pass
    new_rel_ids = set(existing_ids)

    # Clear slide list in out & rebuild as [orig1, res1, orig2, res2, ...]
    for sldId, _ in out_slds:
        out_sldIdLst.remove(sldId)

    pair_count = min(len(orig_slds), len(res_slds))

    # Ensure slides dir exists in out
    slides_dir_out = os.path.join(out_dir, "ppt", "slides")
    if not os.path.exists(slides_dir_out):
        os.makedirs(slides_dir_out, exist_ok=True)

    def next_slide_number():
        existing_slide_files = [
            f for f in os.listdir(slides_dir_out) if f.startswith("slide") and f.endswith(".xml")
        ]
        max_num = 0
        for f in existing_slide_files:
            try:
                num = int(f.replace("slide", "").replace(".xml", ""))
                if num > max_num:
                    max_num = num
            except ValueError:
                pass
        return max_num + 1

    def add_slide_from(source_dir, rel_map, sldId_elem, rId_src, is_morph_target):
        nonlocal new_rel_ids

        src_target = rel_map[rId_src]  # e.g., "slides/slide1.xml"
        src_slide_path = os.path.join(source_dir, "ppt", src_target)

        new_num = next_slide_number()
        new_slide_name = f"slide{new_num}.xml"
        new_target = f"slides/{new_slide_name}"

        dst_slide_path = os.path.join(out_dir, "ppt", new_target)
        shutil.copyfile(src_slide_path, dst_slide_path)

        if is_morph_target:
            add_morph_transition_to_slide(dst_slide_path)

        new_id_numeric = next_id(new_rel_ids)
        new_rel_ids.add(new_id_numeric)
        new_rId = f"rId{new_id_numeric}"

        # Ensure out_rels_root exists
        if out_rels_root is None:
            # create basic relationships root
            rels_root = ET.Element(
                f"{{{NS['r']}}}Relationships"
            )
            out_rels_tree = ET.ElementTree(rels_root)
            globals()["out_rels_root"] = rels_root  # hacky but fine here

        rel_el = ET.SubElement(
            out_rels_root,
            f"{{{NS['r']}}}Relationship",
            {
                "Id": new_rId,
                "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
                "Target": new_target,
            },
        )

        new_sldId = deepcopy(sldId_elem)
        new_sldId.attrib[f"{{{NS['r']}}}id"] = new_rId
        out_sldIdLst.append(new_sldId)

    for idx in range(pair_count):
        orig_sldId_elem, orig_rId = orig_slds[idx]
        res_sldId_elem, res_rId = res_slds[idx]

        add_slide_from(orig_dir, orig_rel_map, orig_sldId_elem, orig_rId, is_morph_target=False)
        add_slide_from(res_dir, res_rel_map, res_sldId_elem, res_rId, is_morph_target=True)

    out_tree.write(out_pres_xml, xml_declaration=True, encoding="utf-8")
    if out_rels_tree is not None:
        out_rels_tree.write(out_rels_path, xml_declaration=True, encoding="utf-8")

    rezip_pptx(out_dir, output_pptx)
    print("Morph deck created at", output_pptx)


if __name__ == "__main__":
    if len(sys.argv) != 4:
        print("Usage: python morph_pptx.py original.pptx resized.pptx output.pptx")
        sys.exit(1)
    main(sys.argv[1], sys.argv[2], sys.argv[3])
