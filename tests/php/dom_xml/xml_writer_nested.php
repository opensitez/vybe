<?php
// vybe-test: php/dom_xml/xml_writer_nested
// origin: languages/php/tests/php/test_dom_xml.rs
// vybe-test-mode: compile

$writer = new XMLWriter();
$writer->openMemory();
$writer->startDocument('1.0', 'UTF-8');
$writer->startElement('catalog');
foreach ([['id' => 1, 'title' => 'Book A'], ['id' => 2, 'title' => 'Book B']] as $book) {
    $writer->startElement('book');
    $writer->writeAttribute('id', $book['id']);
    $writer->writeElement('title', $book['title']);
    $writer->endElement();
}
$writer->endElement();
$writer->endDocument();
$xml = $writer->outputMemory();
$doc = new DOMDocument();
$doc->loadXML($xml);
echo $doc->getElementsByTagName('book')->length;
