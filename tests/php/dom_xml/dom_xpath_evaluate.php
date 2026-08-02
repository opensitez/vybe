<?php
// vybe-test: php/dom_xml/dom_xpath_evaluate
// origin: languages/php/tests/php/test_dom_xml.rs
// vybe-test-mode: compile

$xml = '<data><val>10</val><val>20</val><val>30</val></data>';
$doc = new DOMDocument();
$doc->loadXML($xml);
$xpath = new DOMXPath($doc);
$count = $xpath->evaluate('count(//val)');
echo (int)$count;
