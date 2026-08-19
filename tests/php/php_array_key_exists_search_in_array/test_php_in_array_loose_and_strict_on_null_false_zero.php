<?php
// vybe-test: php/php_array_key_exists_search_in_array/test_php_in_array_loose_and_strict_on_null_false_zero
// origin: languages/php/tests/php/test_php_array_key_exists_search_in_array.rs

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

$a = [0, "0", false, null];
echo (in_array("", $a) ? "empty-loose" : "empty-loose-no") . "|";
echo (in_array("", $a, true) ? "empty-strict" : "empty-strict-no") . "|";
echo (in_array(0, $a, true) ? "zero-strict" : "zero-strict-no");

__vybe_check(ob_get_clean(), "empty-loose|empty-strict-no|zero-strict");
