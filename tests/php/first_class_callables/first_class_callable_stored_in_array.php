<?php
// vybe-test: php/first_class_callables/first_class_callable_stored_in_array
// origin: languages/php/tests/php/test_first_class_callables.rs

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

$ops = [
    'upper' => strtoupper(...),
    'lower' => strtolower(...),
    'rev'   => strrev(...),
];
echo $ops['upper']('hello') . "\n";
echo $ops['lower']('WORLD') . "\n";
echo $ops['rev']('abcde') . "\n";

__vybe_check(ob_get_clean(), "HELLO\nworld\nedcba");
