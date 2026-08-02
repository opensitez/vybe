<?php
// vybe-test: php/php_simplexml_element_attributes_children/test_php_simplexml_dom_import_simplexml
// origin: languages/php/tests/php/test_php_simplexml_element_attributes_children.rs
// vybe-test-mode: compile

$sxe = new SimpleXMLElement("<data><val>123</val></data>");
$dom = dom_import_simplexml($sxe);
echo $dom instanceof DOMElement ? "DOM_IMPORT_OK" : "FAIL";
