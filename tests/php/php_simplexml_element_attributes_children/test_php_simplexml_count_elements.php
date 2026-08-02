<?php
// vybe-test: php/php_simplexml_element_attributes_children/test_php_simplexml_count_elements
// origin: languages/php/tests/php/test_php_simplexml_element_attributes_children.rs
// vybe-test-mode: compile

$xml = '<items><i/><i/><i/></items>';
$sxe = simplexml_load_string($xml);
echo $sxe->count() === 3 ? "SIMPLEXML_COUNT_3_OK" : "FAIL";
