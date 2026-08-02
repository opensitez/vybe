<?php
// vybe-test: php/php_simplexml_element_attributes_children/test_php_simplexml_attributes_iterator
// origin: languages/php/tests/php/test_php_simplexml_element_attributes_children.rs
// vybe-test-mode: compile

$xml = '<tag a="1" b="2"/>';
$sxe = simplexml_load_string($xml);
$attrs = [];
foreach ($sxe->attributes() as $name => $val) {
    $attrs[$name] = (string)$val;
}
echo isset($attrs["a"]) && isset($attrs["b"]) ? "ATTRS_ITERATOR_OK" : "FAIL";
