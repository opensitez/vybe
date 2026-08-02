<?php
// vybe-test: php/dom_xml/dom_create_element
// origin: languages/php/tests/php/test_dom_xml.rs
// vybe-test-mode: compile

$doc = new DOMDocument();
$root = $doc->createElement('root');
$doc->appendChild($root);
echo $doc->documentElement->tagName;
