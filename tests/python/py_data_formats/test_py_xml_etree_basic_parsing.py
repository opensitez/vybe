# vybe-test: python/py_data_formats/test_py_xml_etree_basic_parsing
# origin: languages/python/tests/python/test_py_data_formats.rs

import xml.etree.ElementTree as ET

xml_str = """<catalog>
    <book id="1"><title>Python Cookbook</title><price>39.99</price></book>
    <book id="2"><title>Fluent Python</title><price>49.99</price></book>
</catalog>"""

root = ET.fromstring(xml_str)
print(root.tag)

for book in root.findall("book"):
    title = book.find("title").text
    price = float(book.find("price").text)
    book_id = book.get("id")
    print(f"{book_id}: {title} ${price}")
