<?php
// vybe-test: php/method_chaining/chain_with_interface_return_type
// origin: languages/php/tests/php/test_method_chaining.rs

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

interface Step { public function bump(): self; public function value(): int; }
class Counter implements Step {
    private int $n = 0;
    public function bump(): self { $this->n++; return $this; }
    public function value(): int { return $this->n; }
}
/** @var Step $c */
$c = new Counter();
echo $c->bump()->bump()->value();

__vybe_check(ob_get_clean(), "2");
