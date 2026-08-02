<?php
// vybe-test: php/string_manipulation/str_repeat_boundaries_and_composition_runtime
// origin: languages/php/tests/php/test_string_manipulation.rs

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

echo str_repeat("x", 3);
echo "|";
echo strlen(str_repeat("ab", 0));
echo "|";
echo str_repeat(" ", 2);

__vybe_check(ob_get_clean(), "xxx|0|  ");
