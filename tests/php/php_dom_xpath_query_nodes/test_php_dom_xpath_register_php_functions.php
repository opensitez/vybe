<?php
// vybe-test: php/php_dom_xpath_query_nodes/test_php_dom_xpath_register_php_functions
// origin: languages/php/tests/php/test_php_dom_xpath_query_nodes.rs
// vybe-test-mode: compile

$xml = '<data><name>alice</name></data>';
$doc = new DOMDocument();
$doc->loadXML($xml);
$xpath = new DOMXPath($doc);
if (method_exists($xpath, "registerPhpFunctions")) {
    $xpath->registerPhpFunctions("strtoupper");
    $res = $xpath->query("//name[php:function('strtoupper', string()) = 'ALICE']");
    echo $res->length === 1 ? "PHP_FN_XPATH_OK" : "FAIL";
} else {
    echo "PHP_FN_XPATH_OK";
}
