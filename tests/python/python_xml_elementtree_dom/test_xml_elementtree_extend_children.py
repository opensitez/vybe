# vybe-test: python/python_xml_elementtree_dom/test_xml_elementtree_extend_children
# origin: languages/python/tests/python/test_python_xml_elementtree_dom.rs

import xml.etree.ElementTree as ET
root = ET.Element("root")
children = [ET.Element("child1"), ET.Element("child2")]
root.extend(children)
print([e.tag for e in root])
