<?php
// vybe-test: php/dom_xml/dom_save_xml
// origin: languages/php/tests/php/test_dom_xml.rs
// vybe-test-mode: compile

$doc = new DOMDocument('1.0', 'UTF-8');
$doc->formatOutput = true;
$root = $doc->createElement('data');
$root->appendChild($doc->createTextNode('hello'));
$doc->appendChild($root);
$xml = $doc->saveXML();
echo str_contains($xml, '<data>') ? 'has data tag' : 'missing tag';
echo str_contains($xml, 'hello') ? ':has content' : ':missing content';
