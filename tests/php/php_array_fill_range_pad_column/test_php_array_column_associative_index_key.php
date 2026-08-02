<?php
// vybe-test: php/php_array_fill_range_pad_column/test_php_array_column_associative_index_key
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

$records = [
    ["id" => 2135, "first_name" => "John", "last_name" => "Doe"],
    ["id" => 3245, "first_name" => "Sally", "last_name" => "Smith"],
];
$last_names = array_column($records, "last_name", "id");
echo "2135:" . $last_names[2135] . " 3245:" . $last_names[3245];

__vybe_check(ob_get_clean(), "2135:Doe 3245:Smith");
