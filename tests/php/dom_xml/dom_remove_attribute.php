<?php
// vybe-test: php/dom_xml/dom_remove_attribute
// origin: languages/php/tests/php/test_dom_xml.rs
// vybe-test-mode: compile

$doc = new DOMDocument();
$doc->loadXML('<el id="1" class="x" />');
$el = $doc->documentElement;
$el->removeAttribute('class');
echo $el->hasAttribute('id')    ? 'id ok' : 'id gone';
echo $el->hasAttribute('class') ? ':class still' : ':class removed';
