<?php
// vybe-test: php/php_xml_reader_streaming_parser/test_php_xml_reader_set_parser_property_option
// origin: languages/php/tests/php/test_php_xml_reader_streaming_parser.rs
// vybe-test-mode: compile

$reader = new XMLReader();
$reader->setParserProperty(XMLReader::SUBST_ENTITIES, true);
echo $reader->getParserProperty(XMLReader::SUBST_ENTITIES) ? "PARSER_PROP_OK" : "FAIL";
$reader->close();
