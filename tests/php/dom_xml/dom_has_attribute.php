<?php
// vybe-test: php/dom_xml/dom_has_attribute
// origin: languages/php/tests/php/test_dom_xml.rs
// vybe-test-mode: compile

$doc = new DOMDocument();
$doc->loadXML('<el foo="bar" />');
$el = $doc->documentElement;
echo $el->hasAttribute('foo') ? 'has foo' : 'no foo';
echo $el->hasAttribute('baz') ? 'has baz' : ':no baz';
