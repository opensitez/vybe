<?php
// vybe-test: php/dom_xml/xml_reader_depth
// origin: languages/php/tests/php/test_dom_xml.rs
// vybe-test-mode: compile

$xml = '<a><b><c>deep</c></b></a>';
$reader = new XMLReader();
$reader->XML($xml);
$maxDepth = 0;
while ($reader->read()) {
    if ($reader->depth > $maxDepth) $maxDepth = $reader->depth;
}
$reader->close();
echo $maxDepth;
