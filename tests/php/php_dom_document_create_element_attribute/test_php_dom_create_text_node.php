<?php
// vybe-test: php/php_dom_document_create_element_attribute/test_php_dom_create_text_node
// origin: languages/php/tests/php/test_php_dom_document_create_element_attribute.rs
// vybe-test-mode: compile

$doc = new DOMDocument();
$p = $doc->createElement("p");
$text = $doc->createTextNode("Paragraph content & special <chars>");
$p->appendChild($text);
$doc->appendChild($p);
echo str_contains($doc->saveXML(), "&amp;") || str_contains($doc->saveXML(), "special") ? "TEXT_NODE_OK" : "FAIL";
