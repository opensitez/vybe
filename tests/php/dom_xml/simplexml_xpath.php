<?php
// vybe-test: php/dom_xml/simplexml_xpath
// origin: languages/php/tests/php/test_dom_xml.rs
// vybe-test-mode: compile

$xml = simplexml_load_string('<books><book lang="en"><title>A</title></book><book lang="fr"><title>B</title></book></books>');
$en = $xml->xpath('//book[@lang="en"]/title');
echo count($en) . ':' . $en[0];
