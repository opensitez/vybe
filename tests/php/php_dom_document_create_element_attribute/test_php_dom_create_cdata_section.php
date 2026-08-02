<?php
// vybe-test: php/php_dom_document_create_element_attribute/test_php_dom_create_cdata_section
// origin: languages/php/tests/php/test_php_dom_document_create_element_attribute.rs
// vybe-test-mode: compile

$doc = new DOMDocument();
$cdata = $doc->createCDATASection("<code>if (a < b) {}</code>");
$el = $doc->createElement("script");
$el->appendChild($cdata);
$doc->appendChild($el);
echo str_contains($doc->saveXML(), "CDATA") ? "CDATA_OK" : "FAIL";
