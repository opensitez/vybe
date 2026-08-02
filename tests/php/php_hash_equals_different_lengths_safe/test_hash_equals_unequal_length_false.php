<?php
// vybe-test: php/php_hash_equals_different_lengths_safe/test_hash_equals_unequal_length_false
// origin: languages/php/tests/php/test_php_hash_equals_different_lengths_safe.rs

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

echo hash_equals('short', 'much_longer_string') ? 'equal' : 'not_equal', "\n";

__vybe_check(ob_get_clean(), "not_equal");
