<?php
// vybe-test: php/php_xml_reader_streaming_parser/test_php_xml_reader_next_skips_children
// origin: languages/php/tests/php/test_php_xml_reader_streaming_parser.rs
// vybe-test-mode: compile

$xml = '<list><group><item/></group><target/></list>';
$reader = new XMLReader();
$reader->xml($xml);
$reader->read(); // list
$reader->read(); // group
$reader->next("target"); // Skip group children to target
echo $reader->name === "target" ? "NEXT_TARGET_OK" : "FAIL";
$reader->close();
