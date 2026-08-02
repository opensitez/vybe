<?php
// vybe-test: php/php_simplexml_element_attributes_children/test_php_simplexml_add_attribute
// origin: languages/php/tests/php/test_php_simplexml_element_attributes_children.rs
// vybe-test-mode: compile

$sxe = new SimpleXMLElement("<product/>");
$sxe->addAttribute("price", "19.99");
echo (string)$sxe["price"] === "19.99" ? "ADD_ATTR_OK" : "FAIL";
