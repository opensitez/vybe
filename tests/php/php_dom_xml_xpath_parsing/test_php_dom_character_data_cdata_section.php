<?php
// vybe-test: php/php_dom_xml_xpath_parsing/test_php_dom_character_data_cdata_section
// origin: languages/php/tests/php/test_php_dom_xml_xpath_parsing.rs
// vybe-test-mode: compile

$doc = new DOMDocument();
$cdata = $doc->createCDATASection("<code>if (a < b)</code>");
$root = $doc->createElement("script");
$root->appendChild($cdata);
$doc->appendChild($root);
echo $doc->saveXML();
