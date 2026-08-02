<?php
// vybe-test: php/php_dom_xpath_query_nodes/test_php_dom_xpath_boolean_evaluation
// origin: languages/php/tests/php/test_php_dom_xpath_query_nodes.rs
// vybe-test-mode: compile

$xml = '<data><flag>true</flag></data>';
$doc = new DOMDocument();
$doc->loadXML($xml);
$xpath = new DOMXPath($doc);
$hasFlag = $xpath->evaluate("boolean(//flag)");
echo $hasFlag ? "XPATH_BOOL_TRUE" : "FAIL";
