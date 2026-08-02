<?php
// vybe-test: php/php_dom_document_create_element_attribute/test_php_dom_element_remove_attribute
// origin: languages/php/tests/php/test_php_dom_document_create_element_attribute.rs
// vybe-test-mode: compile

$doc = new DOMDocument();
$el = $doc->createElement("btn");
$el->setAttribute("active", "1");
$el->removeAttribute("active");
echo !$el->hasAttribute("active") ? "REMOVE_ATTR_OK" : "FAIL";
