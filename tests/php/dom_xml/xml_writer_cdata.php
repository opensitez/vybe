<?php
// vybe-test: php/dom_xml/xml_writer_cdata
// origin: languages/php/tests/php/test_dom_xml.rs
// vybe-test-mode: compile

$writer = new XMLWriter();
$writer->openMemory();
$writer->startDocument('1.0');
$writer->startElement('code');
$writer->writeCData('<script>alert("xss")</script>');
$writer->endElement();
$xml = $writer->outputMemory();
echo str_contains($xml, 'CDATA') ? 'has CDATA' : 'no CDATA';
