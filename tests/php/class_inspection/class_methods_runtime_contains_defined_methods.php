<?php
// vybe-test: php/class_inspection/class_methods_runtime_contains_defined_methods
// origin: languages/php/tests/php/test_class_inspection.rs

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
}

$methods = get_class_methods(Calculator::class);
echo in_array('add', $methods) ? 'yes' : 'no';
echo is_string($methods[0]) ? 'string' : 'not';

__vybe_check(ob_get_clean(), "yesstring");
