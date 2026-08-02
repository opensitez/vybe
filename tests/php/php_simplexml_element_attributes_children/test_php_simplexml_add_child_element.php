<?php
// vybe-test: php/php_simplexml_element_attributes_children/test_php_simplexml_add_child_element
// origin: languages/php/tests/php/test_php_simplexml_element_attributes_children.rs
// vybe-test-mode: compile

$sxe = new SimpleXMLElement("<root/>");
$child = $sxe->addChild("setting", "enabled");
echo $sxe->setting == "enabled" ? "ADD_CHILD_OK" : "FAIL";
