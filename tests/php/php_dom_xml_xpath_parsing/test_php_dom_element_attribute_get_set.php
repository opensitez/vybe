<?php
// vybe-test: php/php_dom_xml_xpath_parsing/test_php_dom_element_attribute_get_set
// origin: languages/php/tests/php/test_php_dom_xml_xpath_parsing.rs
// vybe-test-mode: compile

$doc = new DOMDocument();
$elem = $doc->createElement("a", "Click Here");
$elem->setAttribute("href", "https://example.com");
$elem->setAttribute("target", "_blank");

echo $elem->getAttribute("href") . " target=" . $elem->getAttribute("target");
