<?php
// vybe-test: php/string_advanced/strpos_negative_offset_empty_match_runtime
// origin: languages/php/tests/php/test_string_advanced.rs

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

echo strpos("abcabc", "a", -4) === 2 ? "pos2" : "no";
echo "\n";
echo strpos("abc", "") === 0 ? "empty-zero" : "non-zero";

__vybe_check(ob_get_clean(), "no\nempty-zero");
