<?php
// vybe-test: php/method_chaining/chain_with_clone_mid_chain
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

class Draft {
    private int $n = 0;
    public function inc(int $x): static { $this->n += $x; return $this; }
    public function value(): int { return $this->n; }
}
$base = new Draft();
$clone = clone $base;
$base->inc(4);
$clone->inc(1)->inc(2);
echo $base->value() . '|' . $clone->value();

__vybe_check(ob_get_clean(), "4|3");
