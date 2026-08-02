<?php
// vybe-test: php/dom_xml/dom_clone_node
// origin: languages/php/tests/php/test_dom_xml.rs
// vybe-test-mode: compile

$doc = new DOMDocument();
$doc->loadXML('<list><item>hello</item></list>');
$item = $doc->getElementsByTagName('item')->item(0);
$clone = $item->cloneNode(true);
$doc->documentElement->appendChild($clone);
echo $doc->getElementsByTagName('item')->length;
