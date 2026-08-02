<?php
// vybe-test: php/php_functions_arrow_fn_variadic_named/test_php_nested_arrow_functions_currying
// origin: languages/php/tests/php/test_php_functions_arrow_fn_variadic_named.rs

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

$add = fn($x) => fn($y) => $x + $y;
$add5 = $add(5);
echo $add5(10);

__vybe_check(ob_get_clean(), "15");
