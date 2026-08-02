<?php
// vybe-test: php/dom_xml/xml_writer_attributes
// origin: languages/php/tests/php/test_dom_xml.rs
// vybe-test-mode: compile

$writer = new XMLWriter();
$writer->openMemory();
$writer->startDocument('1.0');
$writer->startElement('person');
$writer->writeAttribute('name', 'Alice');
$writer->writeAttribute('age', '30');
$writer->endElement();
$xml = $writer->outputMemory();
echo str_contains($xml, 'name="Alice"') ? 'has name attr' : 'missing';
echo str_contains($xml, 'age="30"')     ? ':has age attr' : ':missing';
