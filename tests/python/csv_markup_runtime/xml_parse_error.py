# vybe-test: python/csv_markup_runtime/xml_parse_error
# origin: languages/python/tests/python/test_csv_markup_runtime.rs

import xml.etree.ElementTree as ET
try:
 ET.fromstring('<unclosed')
 print('ok')
except ET.ParseError:
 print('err')
