<?php
// vybe-test: php/dom_xml/dom_insert_before
// origin: languages/php/tests/php/test_dom_xml.rs
// vybe-test-mode: compile

$doc = new DOMDocument();
$doc->loadXML('<list><b/><c/></list>');
$root = $doc->documentElement;
$a = $doc->createElement('a');
$b = $doc->getElementsByTagName('b')->item(0);
$root->insertBefore($a, $b);
echo $root->firstChild->tagName;
