<?php
// vybe-test: php/php_dom_xpath_query_nodes/test_php_dom_xpath_attribute_predicate_filter
// origin: languages/php/tests/php/test_php_dom_xpath_query_nodes.rs
// vybe-test-mode: compile

$xml = '<users><user status="active">Alice</user><user status="inactive">Bob</user></users>';
$doc = new DOMDocument();
$doc->loadXML($xml);
$xpath = new DOMXPath($doc);
$active = $xpath->query("//user[@status='active']");
echo $active->length === 1 && $active->item(0)->nodeValue === "Alice" ? "XPATH_PREDICATE_OK" : "FAIL";
