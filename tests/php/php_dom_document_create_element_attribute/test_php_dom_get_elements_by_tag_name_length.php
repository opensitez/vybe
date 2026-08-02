<?php
// vybe-test: php/php_dom_document_create_element_attribute/test_php_dom_get_elements_by_tag_name_length
// origin: languages/php/tests/php/test_php_dom_document_create_element_attribute.rs
// vybe-test-mode: compile

$doc = new DOMDocument();
$doc->loadXML("<list><item/> <item/> <item/></list>");
$items = $doc->getElementsByTagName("item");
echo $items->length === 3 ? "TAG_LENGTH_3_OK" : "FAIL";
