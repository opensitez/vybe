<?php
// vybe-test: php/dom_xml/dom_nested_elements
// origin: languages/php/tests/php/test_dom_xml.rs
// vybe-test-mode: compile

$doc = new DOMDocument('1.0', 'UTF-8');
$root = $doc->createElement('catalog');
$doc->appendChild($root);
$book = $doc->createElement('book');
$book->setAttribute('id', '1');
$title = $doc->createElement('title');
$title->appendChild($doc->createTextNode('PHP Manual'));
$book->appendChild($title);
$root->appendChild($book);
echo $root->childNodes->length;
echo ':' . $root->firstChild->getAttribute('id');
