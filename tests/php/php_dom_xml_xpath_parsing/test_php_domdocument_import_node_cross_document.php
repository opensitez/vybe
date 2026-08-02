<?php
// vybe-test: php/php_dom_xml_xpath_parsing/test_php_domdocument_import_node_cross_document
// origin: languages/php/tests/php/test_php_dom_xml_xpath_parsing.rs
// vybe-test-mode: compile

$doc1 = new DOMDocument();
$doc1->loadXML("<root><source>Data</source></root>");

$doc2 = new DOMDocument();
$doc2->loadXML("<target/>");

$imported = $doc2->importNode($doc1->documentElement->firstChild, true);
$doc2->documentElement->appendChild($imported);

echo $doc2->saveXML();
