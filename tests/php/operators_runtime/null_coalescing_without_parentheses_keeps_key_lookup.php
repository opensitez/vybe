<?php
// vybe-test: php/operators_runtime/null_coalescing_without_parentheses_keeps_key_lookup
// origin: languages/php/tests/php/test_operators_runtime.rs

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

$cfg = ['db' => ['host' => null], 'fallback' => '127.0.0.1'];
echo $cfg['db']['host'] ?? $cfg['fallback'];

__vybe_check(ob_get_clean(), "127.0.0.1");
