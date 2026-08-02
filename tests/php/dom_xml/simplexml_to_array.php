<?php
// vybe-test: php/dom_xml/simplexml_to_array
// origin: languages/php/tests/php/test_dom_xml.rs
// vybe-test-mode: compile

$xml = simplexml_load_string('<data><key>value</key><num>42</num></data>');
$arr = json_decode(json_encode($xml), true);
echo $arr['key'] . ':' . $arr['num'];
