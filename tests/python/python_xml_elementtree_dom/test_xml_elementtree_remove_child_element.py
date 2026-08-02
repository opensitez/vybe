# vybe-test: python/python_xml_elementtree_dom/test_xml_elementtree_remove_child_element
# origin: languages/python/tests/python/test_python_xml_elementtree_dom.rs

import xml.etree.ElementTree as ET
root = ET.fromstring("<root><a/><b/><c/></root>")
child_b = root.find("b")
root.remove(child_b)
print([e.tag for e in root])
