<?php
// vybe-test: php/oop/class_name_helpers_runtime
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

class Model {}
$o = new Model();
echo get_class($o), '|';
echo get_class(new Model()), '|';
echo is_a($o, 'Model') ? 'is-a' : 'not';

__vybe_check(ob_get_clean(), "Model|Model|is-a");
