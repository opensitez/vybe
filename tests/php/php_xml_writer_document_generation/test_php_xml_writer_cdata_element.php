<?php
// vybe-test: php/php_xml_writer_document_generation/test_php_xml_writer_cdata_element
// origin: languages/php/tests/php/test_php_xml_writer_document_generation.rs
// vybe-test-mode: compile

$w = new XMLWriter();
$w->openMemory();
$w->startElement("script");
$w->writeCdata("function foo() { return a < b; }");
$w->endElement();
$xml = $w->outputMemory();
echo str_contains($xml, "<![CDATA[") ? "WRITE_CDATA_OK" : "FAIL";
