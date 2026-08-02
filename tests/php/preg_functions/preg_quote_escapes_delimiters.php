<?php
// vybe-test: php/preg_functions/preg_quote_escapes_delimiters
// origin: languages/php/tests/php/test_preg_functions.rs

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

$q = preg_quote('a.b?');
echo str_contains($q, '\\') ? 'quoted' : 'raw';

__vybe_check(ob_get_clean(), "quoted");
