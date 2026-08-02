<?php
// vybe-test: php/classes/class_call_user_func_array_instance_method_runtime
// origin: languages/php/tests/php/test_classes.rs

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

class Adder {
    public function sum(int $a, int $b): int { return $a + $b; }
}
$obj = new Adder();
$callable = [$obj, 'sum'];
echo call_user_func_array($callable, [4, 6]);

__vybe_check(ob_get_clean(), "10");
