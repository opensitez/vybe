<?php
// vybe-test: php/php_dom_document_create_element_attribute/test_php_dom_document_save_html
// origin: languages/php/tests/php/test_php_dom_document_create_element_attribute.rs
// vybe-test-mode: compile

$doc = new DOMDocument();
$doc->loadHTML("<html><body><h1>Title</h1></body></html>");
echo str_contains($doc->saveHTML(), "<h1>Title</h1>") ? "SAVE_HTML_OK" : "FAIL";
