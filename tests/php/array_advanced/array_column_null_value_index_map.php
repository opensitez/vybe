<?php
// vybe-test: php/array_advanced/array_column_null_value_index_map
// origin: languages/php/tests/php/test_array_advanced.rs

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
    ["id" => "u1", "name" => "Alice", "score" => 90],
    ["id" => "u2", "name" => "Bob",   "score" => 85],
];
$indexed = array_column($records, null, "id");
echo $indexed["u1"]["name"];
echo $indexed["u2"]["score"];

__vybe_check(ob_get_clean(), "Alice85");
