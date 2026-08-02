<?php
// vybe-test: php/php_dom_xpath_query_nodes/test_php_dom_xpath_register_namespace_prefix
// origin: languages/php/tests/php/test_php_dom_xpath_query_nodes.rs

function __vybe_check($got, $want) {
    // Match the Rust harness's normalisation: strip \r, then drop trailing
    // newlines (it split on "\n" and popped empty trailing elements).
    $got = str_replace("\r", "", $got);
    $got = rtrim($got, "\n");
    if ($got !== $want) {
        echo "FAIL: want [" . $want . "] got [" . $got . "]\n";
        throw new Exception("assertion failed");
    }
    // Replay the program's own output so running the file by hand still
    // behaves like the program it was extracted from.
    echo $got;
    if ($got !== "") {
        echo "\n";
    }
}

ob_start();

$xml = '<ns:root xmlns:ns="http://example.com/ns"><ns:item>Value</ns:item></ns:root>';
$doc = new DOMDocument();
$doc->loadXML($xml);

$xpath = new DOMXPath($doc);
$xpath->registerNamespace("e", "http://example.com/ns");
$nodes = $xpath->query("//e:item");

echo $nodes->item(0)->nodeValue;

__vybe_check(ob_get_clean(), "Value");
