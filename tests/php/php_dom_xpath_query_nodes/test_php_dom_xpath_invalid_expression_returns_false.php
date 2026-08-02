<?php
// vybe-test: php/php_dom_xpath_query_nodes/test_php_dom_xpath_invalid_expression_returns_false
// origin: languages/php/tests/php/test_php_dom_xpath_query_nodes.rs
// vybe-test-mode: compile

$doc = new DOMDocument();
$doc->loadXML("<root/>");
$xpath = new DOMXPath($doc);
$res = @$xpath->query("///invalid[[[xpath");
echo $res === false ? "INVALID_XPATH_FALSE" : "FAIL";
