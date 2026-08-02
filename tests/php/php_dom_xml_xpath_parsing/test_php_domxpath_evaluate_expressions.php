<?php
// vybe-test: php/php_dom_xml_xpath_parsing/test_php_domxpath_evaluate_expressions
// origin: languages/php/tests/php/test_php_dom_xml_xpath_parsing.rs
// vybe-test-mode: compile

$xml = '<items><item price="10"/><item price="20"/></items>';
$doc = new DOMDocument();
$doc->loadXML($xml);

$xpath = new DOMXPath($doc);
$total = $xpath->evaluate("sum(//item/@price)");
echo "Total: $total";
