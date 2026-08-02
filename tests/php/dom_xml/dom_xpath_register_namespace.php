<?php
// vybe-test: php/dom_xml/dom_xpath_register_namespace
// origin: languages/php/tests/php/test_dom_xml.rs
// vybe-test-mode: compile

$xml = '<root xmlns:app="http://example.com/app"><app:item>value</app:item></root>';
$doc = new DOMDocument();
$doc->loadXML($xml);
$xpath = new DOMXPath($doc);
$xpath->registerNamespace('a', 'http://example.com/app');
$items = $xpath->query('//a:item');
echo $items->length . ':' . $items->item(0)->textContent;
