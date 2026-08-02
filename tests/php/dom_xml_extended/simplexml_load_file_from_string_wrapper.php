<?php
// vybe-test: php/dom_xml_extended/simplexml_load_file_from_string_wrapper
// origin: languages/php/tests/php/test_dom_xml_extended.rs

$xml = simplexml_load_string('<?xml version="1.0"?><root val="1"/>');
echo (string)$xml['val'];
