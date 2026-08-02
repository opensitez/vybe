<?php
// vybe-test: php/dom_xml/dom_load_html
// origin: languages/php/tests/php/test_dom_xml.rs
// vybe-test-mode: compile

$html = '<html><body><h1>Title</h1><p class="intro">Paragraph</p></body></html>';
$doc = new DOMDocument();
@$doc->loadHTML($html);
$h1 = $doc->getElementsByTagName('h1')->item(0);
$p  = $doc->getElementsByTagName('p')->item(0);
echo $h1->textContent . ':' . $p->getAttribute('class');
