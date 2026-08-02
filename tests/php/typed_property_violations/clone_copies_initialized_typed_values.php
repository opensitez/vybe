<?php
// vybe-test: php/typed_property_violations/clone_copies_initialized_typed_values
// origin: languages/php/tests/php/test_typed_property_violations.rs

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

class Cell { public int $v; }
$a = new Cell();
$a->v = 2;
$b = clone $a;
$b->v = 5;
echo $a->v . ':' . $b->v;

__vybe_check(ob_get_clean(), "2:5");
