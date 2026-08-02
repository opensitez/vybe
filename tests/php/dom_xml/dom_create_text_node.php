<?php
// vybe-test: php/dom_xml/dom_create_text_node
// origin: languages/php/tests/php/test_dom_xml.rs
// vybe-test-mode: compile

$doc = new DOMDocument();
$root = $doc->createElement('message');
$text = $doc->createTextNode('Hello, World!');
$root->appendChild($text);
$doc->appendChild($root);
echo $doc->documentElement->textContent;
