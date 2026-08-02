<?php
// vybe-test: php/string_comparison_functions/strcoll_culture_independent_fallback_runtime
// origin: languages/php/tests/php/test_string_comparison_functions.rs

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

setlocale(LC_COLLATE, "C");
$cmp = strcoll("A", "a");
if ($cmp < 0) {
    echo -1;
} else {
    echo 1;
}
echo "|";
echo strcoll("a", "a");

__vybe_check(ob_get_clean(), "-1|0");
