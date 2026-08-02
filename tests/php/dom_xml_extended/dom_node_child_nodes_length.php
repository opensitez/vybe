<?php
// vybe-test: php/dom_xml_extended/dom_node_child_nodes_length
// origin: languages/php/tests/php/test_dom_xml_extended.rs

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
$doc->loadXML('<root><a/><b/></root>');
echo $doc->documentElement->childNodes->length;

__vybe_check(ob_get_clean(), "2");
