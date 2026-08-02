<?php
// vybe-test: php/operators/coalesce_precedence_with_arrays_and_nested_runtime
// origin: languages/php/tests/php/test_operators.rs

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

$cfg = ['a' => ['b' => null], 'fallback' => 'ok'];
echo ($cfg['a']['b'] ?? $cfg['fallback']);
echo '|';
echo (($cfg['a']['b'] ?? null) ?? $cfg['fallback']);
echo '|';
echo (null ?? $cfg['fallback']);
echo '|';
echo (0 ?? $cfg['fallback']);

__vybe_check(ob_get_clean(), "ok|ok|ok|0");
