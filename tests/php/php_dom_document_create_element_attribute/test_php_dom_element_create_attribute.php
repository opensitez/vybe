<?php
// vybe-test: php/php_dom_document_create_element_attribute/test_php_dom_element_create_attribute
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

$doc = new DOMDocument();
$el = $doc->createElement("user");
$attr = $doc->createAttribute("id");
$attr->value = "123";
$el->appendChild($attr);
$doc->appendChild($el);

echo $el->getAttribute("id");

__vybe_check(ob_get_clean(), "123");
