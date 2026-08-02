<?php
// vybe-test: php/php_functions_arrow_first_class_callables/test_php_named_arguments_out_of_order
// origin: languages/php/tests/php/test_php_functions_arrow_first_class_callables.rs

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

function formatUser(string $name, string $role = "guest", bool $active = true) {
    return "$name ($role) active=" . ($active ? "1" : "0");
}

echo formatUser(active: false, name: "Bob");

__vybe_check(ob_get_clean(), "Bob (guest) active=0");
