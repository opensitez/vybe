<?php
// vybe-test: php/class_instantiation/instantiate_final_class_via_child_is_allowed
// origin: languages/php/tests/php/test_class_instantiation.rs

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

final class Leaf {}
echo (new Leaf()) instanceof Leaf ? 'leaf' : 'no';

__vybe_check(ob_get_clean(), "leaf");
