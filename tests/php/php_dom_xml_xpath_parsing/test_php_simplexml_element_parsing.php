<?php
// vybe-test: php/php_dom_xml_xpath_parsing/test_php_simplexml_element_parsing
// origin: languages/php/tests/php/test_php_dom_xml_xpath_parsing.rs
// vybe-test-mode: compile

$xmlStr = "<user><name>Alice</name><email>alice@domain.com</email></user>";
$sxml = simplexml_load_string($xmlStr);
echo "{$sxml->name} <{$sxml->email}>";
