<?php
// vybe-test: php/variable_functions/dynamic_property_method_invocation_runtime
// origin: languages/php/tests/php/test_variable_functions.rs

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

class Greeter {
    public string $method = 'greet';
    public function greet(string $name): string { return "hi $name"; }
}
$g = new Greeter();
$method = $g->method;
echo $g->$method('world');

__vybe_check(ob_get_clean(), "hi world");
