<?php
// vybe-test: php/php_simplexml_element_attributes_children/test_php_simplexml_as_xml_export
// origin: languages/php/tests/php/test_php_simplexml_element_attributes_children.rs
// vybe-test-mode: compile

$sxe = new SimpleXMLElement("<note><to>Tove</to></note>");
$xmlOut = $sxe->asXML();
echo str_contains($xmlOut, "<to>Tove</to>") ? "AS_XML_OK" : "FAIL";
