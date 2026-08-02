<?php
// vybe-test: php/php_xml_reader_streaming_parser/test_php_xml_reader_move_to_attribute_no
// origin: languages/php/tests/php/test_php_xml_reader_streaming_parser.rs
// vybe-test-mode: compile

$xml = '<element a="valA" b="valB"/>';
$reader = new XMLReader();
$reader->xml($xml);
$reader->read();
$reader->moveToAttributeNo(1);
echo $reader->name === "b" && $reader->value === "valB" ? "MOVE_TO_ATTR1_OK" : "FAIL";
$reader->close();
