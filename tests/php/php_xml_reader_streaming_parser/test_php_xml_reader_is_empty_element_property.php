<?php
// vybe-test: php/php_xml_reader_streaming_parser/test_php_xml_reader_is_empty_element_property
// origin: languages/php/tests/php/test_php_xml_reader_streaming_parser.rs
// vybe-test-mode: compile

$xml = '<container><empty/><nonempty>text</nonempty></container>';
$reader = new XMLReader();
$reader->xml($xml);
$reader->read(); // container
$reader->read(); // empty
echo $reader->isEmptyElement ? "IS_EMPTY_ELEMENT_TRUE" : "FAIL";
$reader->close();
