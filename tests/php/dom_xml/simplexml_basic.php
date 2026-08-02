<?php
// vybe-test: php/dom_xml/simplexml_basic
// origin: languages/php/tests/php/test_dom_xml.rs
// vybe-test-mode: compile

$xml = simplexml_load_string('<root><name>Alice</name><age>30</age></root>');
echo $xml->name . ':' . $xml->age;
