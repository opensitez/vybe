<?php
// vybe-test: php/dom_xml/dom_get_elements_by_tag
// origin: languages/php/tests/php/test_dom_xml.rs
// vybe-test-mode: compile

$xml = '<store><book><title>A</title></book><book><title>B</title></book><book><title>C</title></book></store>';
$doc = new DOMDocument();
$doc->loadXML($xml);
$books = $doc->getElementsByTagName('book');
echo $books->length;
