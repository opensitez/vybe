<?php
// vybe-test: php/dom_xml/dom_xpath_query_finds_nodes
// origin: languages/php/tests/php/test_dom_xml.rs

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

$doc = new DOMDocument();
$doc->loadXML('<root><item key="a"/><item key="b"/></root>');
$xpath = new DOMXPath($doc);
echo $xpath->query('//item[@key="b"]')->length;

__vybe_check(ob_get_clean(), "1");
