<?php
// vybe-test: php/dom_xml/dom_xpath_text_content
// origin: languages/php/tests/php/test_dom_xml.rs
// vybe-test-mode: compile

$xml = '<users><user><name>Alice</name><age>30</age></user><user><name>Bob</name><age>25</age></user></users>';
$doc = new DOMDocument();
$doc->loadXML($xml);
$xpath = new DOMXPath($doc);
$names = $xpath->query('//user/name');
$result = [];
foreach ($names as $name) { $result[] = $name->textContent; }
echo implode(',', $result);
