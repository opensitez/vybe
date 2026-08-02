<?php
// vybe-test: php/php_dom_document_create_element_attribute/test_php_dom_document_create_element_and_append
// origin: languages/php/tests/php/test_php_dom_document_create_element_attribute.rs

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

$doc = new DOMDocument("1.0", "UTF-8");
$root = $doc->createElement("root");
$child = $doc->createElement("item", "Hello XML");
$root->appendChild($child);
$doc->appendChild($root);

echo trim($doc->saveXML());

__vybe_check(ob_get_clean(), "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<root><item>Hello XML</item></root>");
