<?php
// vybe-test: php/php_array_key_exists_search_in_array/test_php_array_search_returns_first_matching_key
// origin: languages/php/tests/php/test_php_array_key_exists_search_in_array.rs

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

$arr = ["blue" => 10, "red" => 20, "green" => 10];
$key = array_search(10, $arr);
echo "Key: $key";

__vybe_check(ob_get_clean(), "Key: blue");
