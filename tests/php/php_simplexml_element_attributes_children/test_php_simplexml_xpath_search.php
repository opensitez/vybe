<?php
// vybe-test: php/php_simplexml_element_attributes_children/test_php_simplexml_xpath_search
// origin: languages/php/tests/php/test_php_simplexml_element_attributes_children.rs
// vybe-test-mode: compile

$xml = '<store><book price="10"/><book price="20"/></store>';
$sxe = simplexml_load_string($xml);
$res = $sxe->xpath("//book[@price='20']");
echo count($res) === 1 ? "SIMPLEXML_XPATH_OK" : "FAIL";
