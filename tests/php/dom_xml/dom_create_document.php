<?php
// vybe-test: php/dom_xml/dom_create_document
// origin: languages/php/tests/php/test_dom_xml.rs
// vybe-test-mode: compile

$doc = new DOMDocument('1.0', 'UTF-8');
echo $doc->version . ':' . $doc->encoding;
