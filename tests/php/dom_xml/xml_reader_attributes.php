<?php
// vybe-test: php/dom_xml/xml_reader_attributes
// origin: languages/php/tests/php/test_dom_xml.rs
// vybe-test-mode: compile

$xml = '<items><item id="1" name="A"/><item id="2" name="B"/></items>';
$reader = new XMLReader();
$reader->XML($xml);
$ids = [];
while ($reader->read()) {
    if ($reader->nodeType === XMLReader::ELEMENT && $reader->localName === 'item') {
        $ids[] = $reader->getAttribute('id');
    }
}
$reader->close();
echo implode(',', $ids);
