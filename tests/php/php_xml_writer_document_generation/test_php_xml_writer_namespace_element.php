<?php
// vybe-test: php/php_xml_writer_document_generation/test_php_xml_writer_namespace_element
// origin: languages/php/tests/php/test_php_xml_writer_document_generation.rs
// vybe-test-mode: compile

$w = new XMLWriter();
$w->openMemory();
$w->startElementNs("ns", "element", "http://example.com/ns");
$w->text("Namespace Content");
$w->endElement();
$xml = $w->outputMemory();
echo str_contains($xml, "ns:element") && str_contains($xml, "http://example.com/ns") ? "ELEMENT_NS_OK" : "FAIL";
