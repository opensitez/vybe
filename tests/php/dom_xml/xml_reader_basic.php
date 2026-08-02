<?php
// vybe-test: php/dom_xml/xml_reader_basic
// origin: languages/php/tests/php/test_dom_xml.rs
// vybe-test-mode: compile

$xml = '<?xml version="1.0"?><root><item>one</item><item>two</item></root>';
$reader = new XMLReader();
$reader->XML($xml);
$items = [];
while ($reader->read()) {
    if ($reader->nodeType === XMLReader::ELEMENT && $reader->localName === 'item') {
        $reader->read(); // text node
        $items[] = $reader->value;
    }
}
$reader->close();
echo implode(',', $items);
