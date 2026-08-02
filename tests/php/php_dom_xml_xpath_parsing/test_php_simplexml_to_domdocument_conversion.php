<?php
// vybe-test: php/php_dom_xml_xpath_parsing/test_php_simplexml_to_domdocument_conversion
// origin: languages/php/tests/php/test_php_dom_xml_xpath_parsing.rs
// vybe-test-mode: compile

$sxml = simplexml_load_string("<root><item id='1'/></root>");
$domElem = dom_import_simplexml($sxml);
echo $domElem->nodeName;
