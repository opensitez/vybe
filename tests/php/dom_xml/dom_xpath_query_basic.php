<?php
// vybe-test: php/dom_xml/dom_xpath_query_basic
// origin: languages/php/tests/php/test_dom_xml.rs
// vybe-test-mode: compile

$xml = '<store><book price="10"><title>Alpha</title></book><book price="25"><title>Beta</title></book></store>';
$doc = new DOMDocument();
$doc->loadXML($xml);
$xpath = new DOMXPath($doc);
$books = $xpath->query('//book');
echo $books->length;
