<?php
// vybe-test: php/php_array_transformations_destructuring/test_php_array_keyed_destructuring
// origin: languages/php/tests/php/test_php_array_transformations_destructuring.rs

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

$user = ["name" => "Alice", "role" => "admin", "age" => 30];
["name" => $name, "role" => $role] = $user;
echo "$name is $role";

__vybe_check(ob_get_clean(), "Alice is admin");
