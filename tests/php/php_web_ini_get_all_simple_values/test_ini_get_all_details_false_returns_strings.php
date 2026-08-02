<?php
// vybe-test: php/php_web_ini_get_all_simple_values/test_ini_get_all_details_false_returns_strings
// origin: languages/php/tests/php/test_php_web_ini_get_all_simple_values.rs

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

$all = ini_get_all(null, false);
echo is_array($all) && is_string($all['display_errors'] ?? '') ? 'simple_ini_ok' : 'err', "\n";

__vybe_check(ob_get_clean(), "simple_ini_ok");
