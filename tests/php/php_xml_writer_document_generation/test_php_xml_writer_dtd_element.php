<?php
// vybe-test: php/php_xml_writer_document_generation/test_php_xml_writer_dtd_element
// origin: languages/php/tests/php/test_php_xml_writer_document_generation.rs
// vybe-test-mode: compile

$w = new XMLWriter();
$w->openMemory();
$w->startDtd("html");
$w->endDtd();
$xml = $w->outputMemory();
echo str_contains($xml, "<!DOCTYPE html") ? "WRITE_DTD_OK" : "FAIL";
