<?php
// vybe-test: php/php_simplexml_element_attributes_children/test_php_simplexml_attribute_array_access
// origin: languages/php/tests/php/test_php_simplexml_element_attributes_children.rs

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

$xml = '<item id="42" category="books">PHP Handbook</item>';
$sxe = simplexml_load_string($xml);
echo "ID=" . $sxe["id"] . " Cat=" . $sxe["category"];

__vybe_check(ob_get_clean(), "ID=42 Cat=books");
