<?php
// vybe-test: php/oop_advanced/chaining_with_arrow_and_returned_object
// origin: languages/php/tests/php/test_oop_advanced.rs

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

class Chain {
    public int $value = 1;
    public function inc(int $n): self {
        $this->value += $n;
        return $this;
    }
    public function scale(int $n): self {
        $this->value *= $n;
        return $this;
    }
}
$c = new Chain();
echo $c->inc(2)->scale(3)->value;

__vybe_check(ob_get_clean(), "9");
