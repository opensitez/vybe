<?php
// vybe-test: php/php_dom_document_create_element_attribute/test_php_dom_element_has_attribute_check
// origin: languages/php/tests/php/test_php_dom_document_create_element_attribute.rs
// vybe-test-mode: compile

$doc = new DOMDocument();
$el = $doc->createElement("input");
$el->setAttribute("type", "text");
echo $el->hasAttribute("type") && !$el->hasAttribute("disabled") ? "HAS_ATTR_OK" : "FAIL";
