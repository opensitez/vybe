<?php
// vybe-test: php/php_xml_writer_document_generation/test_php_xml_writer_comment_element
// origin: languages/php/tests/php/test_php_xml_writer_document_generation.rs
// vybe-test-mode: compile

$w = new XMLWriter();
$w->openMemory();
$w->writeComment("This is an XML comment");
$w->startElement("tag");
$w->endElement();
$xml = $w->outputMemory();
echo str_contains($xml, "<!--This is an XML comment-->") ? "WRITE_COMMENT_OK" : "FAIL";
