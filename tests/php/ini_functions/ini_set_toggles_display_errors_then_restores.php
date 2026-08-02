<?php
// vybe-test: php/ini_functions/ini_set_toggles_display_errors_then_restores
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

$old = ini_set('display_errors', '0');
ini_restore('display_errors');
echo ini_get('display_errors') === $old || $old === false ? 'restored' : 'changed';

__vybe_check(ob_get_clean(), "restored");
