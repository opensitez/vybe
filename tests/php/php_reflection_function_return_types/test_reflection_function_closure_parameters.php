<?php
// vybe-test: php/php_reflection_function_return_types/test_reflection_function_closure_parameters
// origin: languages/php/tests/php/test_php_reflection_function_return_types.rs

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

$closure = function(string $name, int $age = 18): string { return "$name:$age"; };
$rf = new ReflectionFunction($closure);
echo $rf->getNumberOfParameters() . ':' . $rf->getNumberOfRequiredParameters(), "\n";

__vybe_check(ob_get_clean(), "2:1");
