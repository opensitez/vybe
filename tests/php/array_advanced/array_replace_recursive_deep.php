<?php
// vybe-test: php/array_advanced/array_replace_recursive_deep
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

$base = ["a" => ["x" => 1, "y" => 2], "b" => 10];
$over = ["a" => ["y" => 99, "z" => 3]];
$result = array_replace_recursive($base, $over);
echo $result["a"]["x"];
echo $result["a"]["y"];
echo $result["a"]["z"];
echo $result["b"];

__vybe_check(ob_get_clean(), "199310");
