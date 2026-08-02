<?php
// vybe-test: php/callables/first_class_callable_instance_method
// origin: languages/php/tests/php/test_callables.rs

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

class Scale { public function __construct(private int $factor) {} public function apply(int $n): int { return $n * $this->factor; } }
$fn = (new Scale(2))->apply(...);
echo $fn(11);

__vybe_check(ob_get_clean(), "22");
