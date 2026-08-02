<?php
// vybe-test: php/php_array_key_exists_search_in_array/test_php_array_key_exists_with_int_and_numeric_string_keys
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

$a = ["0" => "zero", 1 => "one", 1.2 => "float1", true => "bool1"];
echo (array_key_exists(0, $a) ? "0ok" : "0no") . "|";
echo (array_key_exists("1", $a) ? "1ok" : "1no") . "|";
echo (array_key_exists(1, $a) ? "1intok" : "1noint");

__vybe_check(ob_get_clean(), "0ok|1ok|1intok");
