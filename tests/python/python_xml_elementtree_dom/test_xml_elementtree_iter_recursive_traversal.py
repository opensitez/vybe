# vybe-test: python/python_xml_elementtree_dom/test_xml_elementtree_iter_recursive_traversal
# origin: languages/python/tests/python/test_python_xml_elementtree_dom.rs

import xml.etree.ElementTree as ET
xml = "<a><b><c>text1</c></b><c>text2</c></a>"
root = ET.fromstring(xml)
c_tags = [e.text for e in root.iter("c")]
print(c_tags)
