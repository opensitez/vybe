<?php
// vybe-test: php/php_dom_xml_xpath_parsing/test_php_dom_element_attribute_get_set
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

$doc = new DOMDocument();
$elem = $doc->createElement("a", "Click Here");
$elem->setAttribute("href", "https://example.com");
$elem->setAttribute("target", "_blank");

echo $elem->getAttribute("href") . " target=" . $elem->getAttribute("target");


__vybe_check(ob_get_clean(), "https://example.com target=_blank");
