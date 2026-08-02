<?php
// vybe-test: php/php_array_fill_range_pad_column/test_php_array_column_default_value_missing_keys
// origin: languages/php/tests/php/test_php_array_fill_range_pad_column.rs

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

$rows = [
    ["id" => 1, "name" => "A"],
    ["id" => 2],
];
$names = array_column($rows, "name", "id");
echo (array_key_exists(2, $names) ? "has2" : "no2") . "|" . (array_key_exists(3, $names) ? "has3" : "no3");

__vybe_check(ob_get_clean(), "no2|no3");
