<?php
// vybe-test: php/php_xml_reader_streaming_parser/test_php_xml_reader_expand_to_dom_node
// origin: languages/php/tests/php/test_php_xml_reader_streaming_parser.rs
// vybe-test-mode: compile

$xml = '<node attr="val">Content</node>';
$reader = new XMLReader();
$reader->xml($xml);
$reader->read();
$domNode = $reader->expand();
echo $domNode instanceof DOMNode ? "EXPAND_TO_DOM_OK" : "FAIL";
$reader->close();
