<?php
// vybe-test: php/type_juggling_runtime/loose_equal_array_element_against_variable
// origin: languages/php/tests/php/test_type_juggling_runtime.rs

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

$user = ['demo', 'x'];
$name = 'demo';
var_dump($user[0] == $name);

__vybe_check(ob_get_clean(), "bool(true)");
