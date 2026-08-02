# vybe-test: python/python_xml_elementtree_dom/test_xml_elementtree_indent_formatting
# origin: languages/python/tests/python/test_python_xml_elementtree_dom.rs

import xml.etree.ElementTree as ET
root = ET.fromstring("<root><a>1</a><b>2</b></root>")
if hasattr(ET, "indent"):
    ET.indent(root, space="  ")
    xml_str = ET.tostring(root, encoding="unicode")
    print("\n  <a>" in xml_str)
else:
    print(True)
