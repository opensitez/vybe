<?php
// vybe-test: php/string_extra_builtins/stripcslashes_escapes_quotes_backslashes_runtime
// origin: languages/php/tests/php/test_string_extra_builtins.rs

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

echo stripcslashes("a\\nb\\\"c\\'d\\\\e");
echo "|";
echo stripcslashes("line1\\nline2");

__vybe_check(ob_get_clean(), "a\nb\"c'd\\e|line1\nline2");
