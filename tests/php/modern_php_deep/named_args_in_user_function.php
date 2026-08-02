<?php
// vybe-test: php/modern_php_deep/named_args_in_user_function
// origin: languages/php/tests/php/test_modern_php_deep.rs

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

function greet(string $name, string $greeting = "Hello"): string {
    return "$greeting, $name!";
}
echo greet(name: "Alice");
echo greet(name: "Bob", greeting: "Hi");

__vybe_check(ob_get_clean(), "Hello, Alice!Hi, Bob!");
