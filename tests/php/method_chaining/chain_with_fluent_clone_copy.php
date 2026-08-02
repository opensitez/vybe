<?php
// vybe-test: php/method_chaining/chain_with_fluent_clone_copy
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

class Token {
    public function __construct(private int $n = 0) {}
    public function add(int $d): static { $this->n += $d; return $this; }
    public function value(): int { return $this->n; }
}
$base = new Token(1);
$next = clone $base;
echo $base->add(2)->value() . '|' . $next->add(3)->value();

__vybe_check(ob_get_clean(), "3|4");
