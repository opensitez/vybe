<?php
// vybe-test: php/oop/clone_reinitializes_with_magic_clone_runtime
// origin: languages/php/tests/php/test_oop.rs

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
    public array $values;
    public function __construct() { $this->values = [1, 2, 3]; }
    public function __clone(): void { $this->values[] = 9; }
}
$a = new Counter();
$b = clone $a;
$a->values[0] = 9;
echo implode(',', $b->values);

__vybe_check(ob_get_clean(), "1,2,3,9");
