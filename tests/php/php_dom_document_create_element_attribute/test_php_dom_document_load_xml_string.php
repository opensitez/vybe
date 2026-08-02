<?php
// vybe-test: php/php_dom_document_create_element_attribute/test_php_dom_document_load_xml_string
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

$xml = "<config><setting name='debug'>true</setting></config>";
$doc = new DOMDocument();
$doc->loadXML($xml);

$nodes = $doc->getElementsByTagName("setting");
echo $nodes->item(0)->nodeValue;

__vybe_check(ob_get_clean(), "true");
