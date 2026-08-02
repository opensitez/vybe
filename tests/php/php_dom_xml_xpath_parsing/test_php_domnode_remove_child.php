<?php
// vybe-test: php/php_dom_xml_xpath_parsing/test_php_domnode_remove_child
// origin: languages/php/tests/php/test_php_dom_xml_xpath_parsing.rs
// vybe-test-mode: compile

$doc = new DOMDocument();
$doc->loadXML("<root><child1/><child2/></root>");
$root = $doc->documentElement;
$root->removeChild($root->firstChild);
echo $doc->saveXML();
