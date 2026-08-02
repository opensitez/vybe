# vybe-test: python/python_xml_elementtree_dom/test_xml_elementtree_insert_child_element
# origin: languages/python/tests/python/test_python_xml_elementtree_dom.rs

import xml.etree.ElementTree as ET
root = ET.fromstring("<root><a/><c/></root>")
b = ET.Element("b")
root.insert(1, b)
print([e.tag for e in root])
