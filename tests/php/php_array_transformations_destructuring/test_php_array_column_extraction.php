<?php
// vybe-test: php/php_array_transformations_destructuring/test_php_array_column_extraction
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

$records = [
    ["id" => 101, "name" => "Alice"],
    ["id" => 102, "name" => "Bob"],
    ["id" => 103, "name" => "Charlie"],
];
$names = array_column($records, "name", "id");
echo $names[102];

__vybe_check(ob_get_clean(), "Bob");
