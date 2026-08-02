<?php
// vybe-test: php/php_xml_writer_document_generation/test_php_xml_writer_pi_processing_instruction
// origin: languages/php/tests/php/test_php_xml_writer_document_generation.rs
// vybe-test-mode: compile

$w = new XMLWriter();
$w->openMemory();
$w->writePi("php-stylesheet", 'href="style.css"');
$w->startElement("page");
$w->endElement();
$xml = $w->outputMemory();
echo str_contains($xml, "<?php-stylesheet") ? "WRITE_PI_OK" : "FAIL";
