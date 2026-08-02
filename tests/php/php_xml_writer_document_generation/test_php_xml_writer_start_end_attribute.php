<?php
// vybe-test: php/php_xml_writer_document_generation/test_php_xml_writer_start_end_attribute
// origin: languages/php/tests/php/test_php_xml_writer_document_generation.rs
// vybe-test-mode: compile

$w = new XMLWriter();
$w->openMemory();
$w->startElement("link");
$w->startAttribute("href");
$w->text("https://example.com");
$w->endAttribute();
$w->endElement();
$xml = $w->outputMemory();
echo str_contains($xml, 'href="https://example.com"') ? "START_END_ATTR_OK" : "FAIL";
