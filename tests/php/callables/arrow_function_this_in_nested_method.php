<?php
// vybe-test: php/callables/arrow_function_this_in_nested_method
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

class Outer {
    private int $n = 4;
    public function wrap(): int {
        $inner = fn(): int => $this->n * 2;
        return $inner();
    }
}
echo (new Outer())->wrap();

__vybe_check(ob_get_clean(), "8");
