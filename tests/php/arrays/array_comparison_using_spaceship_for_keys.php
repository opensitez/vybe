<?php
// vybe-test: php/arrays/array_comparison_using_spaceship_for_keys
// origin: languages/php/tests/php/test_arrays.rs

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

$a = ['x' => 1, 'y' => 2];
$b = ['x' => 1, 'y' => 3];
echo ($a == $b ? 'eq' : 'neq') . '|';
echo ($a <=> $b);

__vybe_check(ob_get_clean(), "neq|-1");
