<?php
// vybe-test: php/php_string_searching_substring_positions/test_php_strrpos_case_sensitive_reverse_search
// origin: languages/php/tests/php/test_php_string_searching_substring_positions.rs

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

$text = "The quick brown fox jumps over the lazy dog";
$pos = strrpos($text, "the");
echo "Last 'the' at offset: $pos";

__vybe_check(ob_get_clean(), "Last 'the' at offset: 31");
