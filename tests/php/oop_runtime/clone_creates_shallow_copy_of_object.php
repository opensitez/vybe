<?php
// vybe-test: php/oop_runtime/clone_creates_shallow_copy_of_object
// origin: languages/php/tests/php/test_oop_runtime.rs

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

class Box { public function __construct(public int $n) {} }
$a = new Box(1);
$b = clone $a;
$b->n = 2;
echo $a->n . $b->n;

__vybe_check(ob_get_clean(), "12");
