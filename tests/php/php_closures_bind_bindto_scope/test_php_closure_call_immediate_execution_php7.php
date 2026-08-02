<?php
// vybe-test: php/php_closures_bind_bindto_scope/test_php_closure_call_immediate_execution_php7
// origin: languages/php/tests/php/test_php_closures_bind_bindto_scope.rs

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

class Person {
    private string $name = "Charlie";
}

$p = new Person();
$greeting = (function(string $prefix) {
    return "$prefix {$this->name}";
})->call($p, "Hello");

echo $greeting;

__vybe_check(ob_get_clean(), "Hello Charlie");
