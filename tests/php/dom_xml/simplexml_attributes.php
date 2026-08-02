<?php
// vybe-test: php/dom_xml/simplexml_attributes
// origin: languages/php/tests/php/test_dom_xml.rs
// vybe-test-mode: compile

$xml = simplexml_load_string('<user id="42" role="admin"><name>Bob</name></user>');
echo $xml['id'] . ':' . $xml['role'] . ':' . $xml->name;
