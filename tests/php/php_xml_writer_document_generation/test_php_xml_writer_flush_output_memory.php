<?php
// vybe-test: php/php_xml_writer_document_generation/test_php_xml_writer_flush_output_memory
// origin: languages/php/tests/php/test_php_xml_writer_document_generation.rs
// vybe-test-mode: compile

$w = new XMLWriter();
$w->openMemory();
$w->writeElement("a", "b");
$chunk1 = $w->flush();
$w->writeElement("c", "d");
$chunk2 = $w->flush();
echo str_contains($chunk1, "<a>b</a>") && str_contains($chunk2, "<c>d</c>") ? "FLUSH_MEMORY_OK" : "FAIL";
