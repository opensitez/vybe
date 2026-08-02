<?php
// vybe-test: php/method_chaining/chain_operator_precedence_with_parentheses
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

class Counter {
    private int $n = 0;
    public function inc(int $d): static { $this->n += $d; return $this; }
    public function value(): int { return $this->n; }
}
echo ((new Counter())->inc(1)->value() + 1) . '|';
echo ((new Counter())->inc(2)->value()) + (1 + 1);

__vybe_check(ob_get_clean(), "2|4");
