<?php
// vybe-test: php/object_comparison/object_equality_compares_all_properties
// origin: languages/php/tests/php/test_object_comparison.rs

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

class Config { public string $a; public int $b; }
$x = new Config(); $x->a = 'hello'; $x->b = 1;
$y = new Config(); $y->a = 'hello'; $y->b = 1;
$z = new Config(); $z->a = 'hello'; $z->b = 2;
echo ($x == $y ? 'eq' : 'ne') . ',' . ($x == $z ? 'eq' : 'ne');

__vybe_check(ob_get_clean(), "eq,ne");
