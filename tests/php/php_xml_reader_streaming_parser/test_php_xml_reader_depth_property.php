<?php
// vybe-test: php/php_xml_reader_streaming_parser/test_php_xml_reader_depth_property
// origin: languages/php/tests/php/test_php_xml_reader_streaming_parser.rs
// vybe-test-mode: compile

$xml = '<level0><level1><level2/></level1></level0>';
$reader = new XMLReader();
$reader->xml($xml);
$reader->read(); // level0
$reader->read(); // level1
echo $reader->depth === 1 ? "DEPTH_1_OK" : "FAIL";
$reader->close();
