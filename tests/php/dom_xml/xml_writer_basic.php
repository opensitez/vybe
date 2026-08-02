<?php
// vybe-test: php/dom_xml/xml_writer_basic
// origin: languages/php/tests/php/test_dom_xml.rs
// vybe-test-mode: compile

$writer = new XMLWriter();
$writer->openMemory();
$writer->startDocument('1.0', 'UTF-8');
$writer->startElement('root');
$writer->writeElement('child', 'value');
$writer->endElement();
$writer->endDocument();
$xml = $writer->outputMemory();
echo str_contains($xml, '<root>') ? 'has root' : 'no root';
echo str_contains($xml, 'value') ? ':has value' : ':no value';
