# vybe-test: python/py_data_formats/test_py_xml_etree_xpath_find
# origin: languages/python/tests/python/test_py_data_formats.rs

import xml.etree.ElementTree as ET

xml_str = """<root>
    <item type="fruit">apple</item>
    <item type="veg">carrot</item>
    <item type="fruit">banana</item>
</root>"""

root = ET.fromstring(xml_str)
fruits = root.findall(".//item[@type='fruit']")
print([el.text for el in fruits])
