# vybe-test: python/python_xml_elementtree_dom/test_xml_elementtree_find_and_findall
# origin: languages/python/tests/python/test_python_xml_elementtree_dom.rs

import xml.etree.ElementTree as ET
xml = """<catalog>
    <book id="bk101"><title>Python Guide</title></book>
    <book id="bk102"><title>Rust Guide</title></book>
</catalog>"""
root = ET.fromstring(xml)
first_book = root.find("book")
print(first_book.attrib["id"])

all_titles = [t.text for t in root.findall("book/title")]
print(all_titles)
