<?php
// vybe-test: php/dom_xml/dom_load_xml_string
// origin: languages/php/tests/php/test_dom_xml.rs
// vybe-test-mode: compile

$xml = '<?xml version="1.0"?><root><item id="1">First</item><item id="2">Second</item></root>';
$doc = new DOMDocument();
$doc->loadXML($xml);
$items = $doc->getElementsByTagName('item');
echo $items->length;
echo ':' . $items->item(0)->textContent;
