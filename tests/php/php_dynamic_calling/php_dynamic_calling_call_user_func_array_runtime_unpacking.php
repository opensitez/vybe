<?php
// vybe-test: php/php_dynamic_calling/php_dynamic_calling_call_user_func_array_runtime_unpacking
// origin: languages/php/tests/php/test_php_dynamic_calling.rs

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

function combine(string $a, string $b, string $c): string {
    return $a . '-' . $b . '-' . $c;
}
$fn = 'combine';
echo call_user_func_array($fn, ['a', 'b', 'c']);

__vybe_check(ob_get_clean(), "a-b-c");
