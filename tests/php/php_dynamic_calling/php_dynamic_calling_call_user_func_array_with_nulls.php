<?php
// vybe-test: php/php_dynamic_calling/php_dynamic_calling_call_user_func_array_with_nulls
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

function combine_three(string $a, ?string $b = null, string $c = 'c'): string {
    return $a . '|' . ($b ?? 'none') . '|' . $c;
}

echo call_user_func_array('combine_three', ['a', null, 'z']);

__vybe_check(ob_get_clean(), "a|none|z");
