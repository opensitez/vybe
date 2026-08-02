<?php
// vybe-test: php/php_dom_xml_xpath_parsing/test_php_domdocument_load_xml_and_xpath_query
// origin: languages/php/tests/php/test_php_dom_xml_xpath_parsing.rs

$xml = <<<XML
<?xml version="1.0"?>
<catalog>
    <book id="bk101"><title>PHP 8 in Action</title><price>39.95</price></book>
    <book id="bk102"><title>Rust Systems</title><price>49.95</price></book>
</catalog>
XML;

$doc = new DOMDocument();
$doc->loadXML($xml);

$xpath = new DOMXPath($doc);
$titles = $xpath->query("//book/title");

$out = [];
foreach ($titles as $node) {
    $out[] = $node->nodeValue;
}
echo implode(" | ", $out);
