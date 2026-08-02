<?php
// vybe-test: php/php_dom_xpath_query_nodes/test_php_dom_xpath_context_node_query
// origin: languages/php/tests/php/test_php_dom_xpath_query_nodes.rs
// vybe-test-mode: compile

$xml = '<section><group id="g1"><item>1</item></group></section>';
$doc = new DOMDocument();
$doc->loadXML($xml);
$xpath = new DOMXPath($doc);
$group = $xpath->query("//group")->item(0);
$item = $xpath->query("./item", $group);
echo $item->item(0)->nodeValue === "1" ? "CONTEXT_NODE_OK" : "FAIL";
