<?php
// vybe-test: php/array_replace_recursive_deep/array_replace_recursive_with_numeric_string_keys
// origin: languages/php/tests/php/test_array_replace_recursive_deep.rs

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

$base = ["1" => ["a" => 1], "2" => ["b" => 2]];
$patch = [1 => ["a" => 9], 2 => ["c" => 3]];
$res = array_replace_recursive($base, $patch);
echo $res[1]["a"] . "|" . $res[2]["b"] . "|" . $res[2]["c"];

__vybe_check(ob_get_clean(), "9|2|3");
