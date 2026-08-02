<?php
// vybe-test: php/dom_xml/simplexml_children
// origin: languages/php/tests/php/test_dom_xml.rs
// vybe-test-mode: compile

$xml = simplexml_load_string('<list><item>a</item><item>b</item><item>c</item></list>');
$count = 0;
foreach ($xml->item as $item) { $count++; }
echo $count;
echo ':' . $xml->item[1];
