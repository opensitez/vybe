<?php
// vybe-test: php/array_advanced/array_splice_string_offset_and_preserve_values
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

$data = ["a" => 1, "b" => 2, "c" => 3, "d" => 4];
$removed = array_splice($data, "1", 2, [9, 10]);
echo implode(",", $data);
echo "|";
echo implode(",", $removed);

__vybe_check(ob_get_clean(), "1,9,10,4|2,3");
