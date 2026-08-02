<?php
// vybe-test: php/php_xml_reader_streaming_parser/test_php_xml_reader_read_inner_xml_string
// origin: languages/php/tests/php/test_php_xml_reader_streaming_parser.rs
// vybe-test-mode: compile

$xml = '<wrapper><content>Hello XMLReader</content></wrapper>';
$reader = new XMLReader();
$reader->xml($xml);
$reader->read(); // wrapper
echo str_contains($reader->readInnerXML(), "<content>") ? "INNER_XML_OK" : "FAIL";
$reader->close();
