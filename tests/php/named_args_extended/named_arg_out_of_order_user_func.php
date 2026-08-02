<?php
// vybe-test: php/named_args_extended/named_arg_out_of_order_user_func
// origin: languages/php/tests/php/test_named_args_extended.rs

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

function greet(string $name, string $greeting = 'Hello'): string {
    return "$greeting, $name!";
}
echo greet(greeting: 'Hi', name: 'Alice');

__vybe_check(ob_get_clean(), "Hi, Alice!");
