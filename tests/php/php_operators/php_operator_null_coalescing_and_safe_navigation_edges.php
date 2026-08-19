<?php
// vybe-test: php/php_operators/php_operator_null_coalescing_and_safe_navigation_edges
// origin: languages/php/tests/php/test_php_operators.rs

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

echo (function() { $value = null; $value ??= 'default'; return $value; })(), "\n";
echo (function() { $value = 0; $value ??= 'default'; return $value; })(), "\n";
echo (function() { $user = null; return $user?->name ?? 'anon'; })(), "\n";
echo (function() { $user = (object)['name' => 'Ada']; return $user?->name ?? 'anon'; })(), "\n";


__vybe_check(ob_get_clean(), "default\n0\nanon\nAda");
