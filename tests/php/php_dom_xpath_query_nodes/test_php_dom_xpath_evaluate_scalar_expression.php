<?php
// vybe-test: php/php_dom_xpath_query_nodes/test_php_dom_xpath_evaluate_scalar_expression
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

$xml = '<items><item price="10"/><item price="20"/></items>';
$doc = new DOMDocument();
$doc->loadXML($xml);

$xpath = new DOMXPath($doc);
$count = $xpath->evaluate("count(//item)");
echo "Count: $count";

__vybe_check(ob_get_clean(), "Count: 2");
