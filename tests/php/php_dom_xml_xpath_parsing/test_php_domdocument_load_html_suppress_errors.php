<?php
// vybe-test: php/php_dom_xml_xpath_parsing/test_php_domdocument_load_html_suppress_errors
// origin: languages/php/tests/php/test_php_dom_xml_xpath_parsing.rs
// vybe-test-mode: compile

$html = '<div class="content"><p>Unclosed paragraph</div>';
$doc = new DOMDocument();
libxml_use_internal_errors(true);
$doc->loadHTML($html);
libxml_clear_errors();

$p = $doc->getElementsByTagName("p")->item(0);
echo $p ? $p->textContent : "NO_P";
