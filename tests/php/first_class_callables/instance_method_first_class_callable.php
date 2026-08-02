<?php
// vybe-test: php/first_class_callables/instance_method_first_class_callable
// origin: languages/php/tests/php/test_first_class_callables.rs

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

class Multiplier {
    public function __construct(private int $factor) {}
    public function multiply(int $n): int { return $n * $this->factor; }
}
$m = new Multiplier(3);
$fn = $m->multiply(...);
echo $fn(4) . "\n";
$result = array_map($fn, [1, 2, 3]);
echo implode(',', $result) . "\n";

__vybe_check(ob_get_clean(), "12\n3,6,9");
