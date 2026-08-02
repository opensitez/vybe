<?php
// vybe-test: php/array_advanced/array_slice_offsets_and_flags
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
$tail = array_slice($data, -3, 2, true);
echo implode(",", array_keys($tail)) . "|" . implode(",", $tail);
echo "|";
$values = array_slice($data, 1, 2, false);
echo implode("", $values);

__vybe_check(ob_get_clean(), "b,c|2,3|23");
