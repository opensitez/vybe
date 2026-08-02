<?php
// vybe-test: php/callables/bind_then_invoke_in_loop
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

class Counter { private int $n = 0; public function bump(): int { return ++$this->n; } }
$inc = Closure::bind(function(): int { return $this->bump(); }, new Counter(), Counter::class);
echo $inc() . $inc();

__vybe_check(ob_get_clean(), "12");
