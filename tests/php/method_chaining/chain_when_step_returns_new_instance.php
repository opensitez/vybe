<?php
// vybe-test: php/method_chaining/chain_when_step_returns_new_instance
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

class Numberer {
    private int $v;
    public function __construct(int $v = 0) { $this->v = $v; }
    public function withOffset(int $d): Numberer { return new Numberer($this->v + $d); }
    public function value(): int { return $this->v; }
}
$n = (new Numberer())->withOffset(4)->withOffset(5)->value();
echo $n;

__vybe_check(ob_get_clean(), "9");
