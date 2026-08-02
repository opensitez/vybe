<?php
// vybe-test: php/php_dom_xpath_query_nodes/test_php_dom_xpath_query_returns_nodelist
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

$xml = '<catalog><book id="b1"><title>PHP 8</title></book><book id="b2"><title>Rust</title></book></catalog>';
$doc = new DOMDocument();
$doc->loadXML($xml);

$xpath = new DOMXPath($doc);
$titles = $xpath->query("//book/title");

$out = [];
foreach ($titles as $node) {
    $out[] = $node->nodeValue;
}
echo implode(",", $out);

__vybe_check(ob_get_clean(), "PHP 8,Rust");
