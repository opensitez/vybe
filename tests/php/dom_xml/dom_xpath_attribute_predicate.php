<?php
// vybe-test: php/dom_xml/dom_xpath_attribute_predicate
// origin: languages/php/tests/php/test_dom_xml.rs
// vybe-test-mode: compile

$xml = '<items><item id="1" active="true"/><item id="2" active="false"/><item id="3" active="true"/></items>';
$doc = new DOMDocument();
$doc->loadXML($xml);
$xpath = new DOMXPath($doc);
$active = $xpath->query('//item[@active="true"]');
echo $active->length;
