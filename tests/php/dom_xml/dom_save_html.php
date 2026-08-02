<?php
// vybe-test: php/dom_xml/dom_save_html
// origin: languages/php/tests/php/test_dom_xml.rs
// vybe-test-mode: compile

$doc = new DOMDocument();
@$doc->loadHTML('<html><body><p>Test</p></body></html>');
$html = $doc->saveHTML();
echo str_contains($html, '<p>Test</p>') ? 'ok' : 'fail';
