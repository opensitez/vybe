<?php
// vybe-test: php/dom_xml/dom_remove_child
// origin: languages/php/tests/php/test_dom_xml.rs
// vybe-test-mode: compile

$doc = new DOMDocument();
$doc->loadXML('<list><a/><b/><c/></list>');
$root = $doc->documentElement;
$b = $doc->getElementsByTagName('b')->item(0);
$root->removeChild($b);
echo $root->childNodes->length;
