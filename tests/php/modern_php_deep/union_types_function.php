<?php
// vybe-test: php/modern_php_deep/union_types_function
// origin: languages/php/tests/php/test_modern_php_deep.rs

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

function stringify(int|float|string $val): string {
    if (is_int($val)) return "int:$val";
    if (is_float($val)) return "float:$val";
    return "str:$val";
}
echo stringify(42);
echo stringify(3.14);
echo stringify("hello");

__vybe_check(ob_get_clean(), "int:42float:3.14str:hello");
