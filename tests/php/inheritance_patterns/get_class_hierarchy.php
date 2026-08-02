<?php
// vybe-test: php/inheritance_patterns/get_class_hierarchy
// origin: languages/php/tests/php/test_inheritance_patterns.rs

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

class A {} class B extends A {} class C extends B {}
$c = new C;
echo get_class($c) . ':' . get_parent_class($c) . ':' . is_a($c, 'A') ? 'a' : 'not', "\n";

__vybe_check(ob_get_clean(), "a");
