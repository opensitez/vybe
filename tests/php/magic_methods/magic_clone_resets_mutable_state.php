<?php
// vybe-test: php/magic_methods/magic_clone_resets_mutable_state
// origin: languages/php/tests/php/test_magic_methods.rs

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
    public int $count = 0;
    public function increment(): void { $this->count++; }
    public function __clone() { $this->count = 0; }
}
$a = new Counter();
$a->increment();
$a->increment();
$b = clone $a;
echo $a->count;
echo $b->count;

__vybe_check(ob_get_clean(), "20");
