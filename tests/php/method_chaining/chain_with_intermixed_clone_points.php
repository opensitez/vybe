<?php
// vybe-test: php/method_chaining/chain_with_intermixed_clone_points
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
    public function __construct(private int $v) {}
    public function with(int $d): Counter { return new Counter($this->v + $d); }
    public function plus(int $d): self { $this->v += $d; return $this; }
    public function value(): int { return $this->v; }
}
$base = new Counter(1);
$copy = $base->with(4)->plus(3);
echo $base->value() . '|' . $copy->value();

__vybe_check(ob_get_clean(), "1|8");
