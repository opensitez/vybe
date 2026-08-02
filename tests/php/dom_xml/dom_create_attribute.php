<?php
// vybe-test: php/dom_xml/dom_create_attribute
// origin: languages/php/tests/php/test_dom_xml.rs
// vybe-test-mode: compile

$doc = new DOMDocument();
$el = $doc->createElement('person');
$el->setAttribute('name', 'Alice');
$el->setAttribute('age', '30');
$doc->appendChild($el);
echo $el->getAttribute('name') . ':' . $el->getAttribute('age');
