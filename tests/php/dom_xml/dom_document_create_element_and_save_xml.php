<?php
// vybe-test: php/dom_xml/dom_document_create_element_and_save_xml
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

$doc = new DOMDocument('1.0', 'UTF-8');
$root = $doc->createElement('root');
$doc->appendChild($root);
$child = $doc->createElement('name', 'vybe');
$root->appendChild($child);
echo $doc->saveXML($child);

__vybe_check(ob_get_clean(), "<name>vybe</name>");
