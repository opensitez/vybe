<?php
// vybe-test: php/php_dom_xml_xpath_parsing/test_php_domdocument_create_element_and_append
// origin: languages/php/tests/php/test_php_dom_xml_xpath_parsing.rs

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
$root = $doc->createElement("response");
$status = $doc->createElement("status", "success");
$root->appendChild($status);
$doc->appendChild($root);

echo $doc->saveXML();

__vybe_check(ob_get_clean(), "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<response><status>success</status></response>");
