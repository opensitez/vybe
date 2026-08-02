<?php
// vybe-test: php/php_dom_xpath_query_nodes/test_php_dom_xpath_string_evaluation
// origin: languages/php/tests/php/test_php_dom_xpath_query_nodes.rs
// vybe-test-mode: compile

$xml = '<root><val>TestString</val></root>';
$doc = new DOMDocument();
$doc->loadXML($xml);
$xpath = new DOMXPath($doc);
$str = $xpath->evaluate("string(//val)");
echo $str === "TestString" ? "EVALUATE_STRING_OK" : "FAIL";
