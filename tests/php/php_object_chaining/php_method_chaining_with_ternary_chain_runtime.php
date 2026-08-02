<?php
// vybe-test: php/php_object_chaining/php_method_chaining_with_ternary_chain_runtime
// origin: languages/php/tests/php/test_php_object_chaining.rs

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
    public function __construct(public int $value) {}
    public function inc(int $n): self { $this->value += $n; return $this; }
    public function dec(int $n): self { $this->value -= $n; return $this; }
    public function done(): int { return $this->value; }
}
$c = (new Calculator(3))->inc(7)->dec(4);
echo $c->done();

__vybe_check(ob_get_clean(), "6");
