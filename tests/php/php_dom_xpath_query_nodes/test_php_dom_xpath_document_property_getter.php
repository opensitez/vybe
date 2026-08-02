<?php
// vybe-test: php/php_dom_xpath_query_nodes/test_php_dom_xpath_document_property_getter
// origin: languages/php/tests/php/test_php_dom_xpath_query_nodes.rs
// vybe-test-mode: compile

$doc = new DOMDocument();
$doc->loadXML("<root/>");
$xpath = new DOMXPath($doc);
echo $xpath->document === $doc ? "XPATH_DOC_PROP_OK" : "FAIL";
