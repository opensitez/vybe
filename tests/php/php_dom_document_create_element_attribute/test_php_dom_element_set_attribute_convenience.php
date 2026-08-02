<?php
// vybe-test: php/php_dom_document_create_element_attribute/test_php_dom_element_set_attribute_convenience
// origin: languages/php/tests/php/test_php_dom_document_create_element_attribute.rs
// vybe-test-mode: compile

$doc = new DOMDocument();
$el = $doc->createElement("div");
$el->setAttribute("class", "container");
echo $el->getAttribute("class") === "container" ? "SET_ATTR_OK" : "FAIL";
