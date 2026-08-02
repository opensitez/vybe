<?php
// vybe-test: php/object_model/invoke_used_as_callback
// origin: languages/php/tests/php/test_object_model.rs

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

class Adder { public function __construct(private int $n) {} public function __invoke(int $x): int { return $x + $this->n; } }
echo implode(',', array_map(new Adder(10), [1,2,3]));

__vybe_check(ob_get_clean(), "11,12,13");
