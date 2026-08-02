<?php
// vybe-test: php/php_array_key_exists_search_in_array/test_php_array_keys_with_object_values
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

$o1 = (object)["id" => 1];
$o2 = (object)["id" => 2];
$a = ["a" => $o1, "b" => $o2];
$keys = array_keys($a);
echo implode(",", $keys);

__vybe_check(ob_get_clean(), "a,b");
