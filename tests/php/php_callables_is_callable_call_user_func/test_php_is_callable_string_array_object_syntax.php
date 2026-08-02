<?php
// vybe-test: php/php_callables_is_callable_call_user_func/test_php_is_callable_string_array_object_syntax
// origin: languages/php/tests/php/test_php_callables_is_callable_call_user_func.rs

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

class Calculator {
    public function add(int $a, int $b): int { return $a + $b; }
    public static function multiply(int $a, int $b): int { return $a * $b; }
}

$c = new Calculator();

echo is_callable("strlen") ? "1" : "0";
echo is_callable([$c, "add"]) ? "1" : "0";
echo is_callable([Calculator::class, "multiply"]) ? "1" : "0";
echo is_callable("Calculator::multiply") ? "1" : "0";

__vybe_check(ob_get_clean(), "1111");
