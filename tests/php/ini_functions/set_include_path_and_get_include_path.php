<?php
// vybe-test: php/ini_functions/set_include_path_and_get_include_path
// origin: languages/php/tests/php/test_ini_functions.rs

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

$old = set_include_path('/tmp/custom_include');
echo get_include_path() === '/tmp/custom_include' ? 'include_path_set' : 'err';

__vybe_check(ob_get_clean(), "include_path_set");
