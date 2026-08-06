<?php
// vybe-test: php/array_advanced/array_combine_with_numeric_string_keys
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

$keys = ["10", 20, "30.0"];
$vals = ["a", "b", "c"];
$combined = array_combine($keys, $vals);
echo $combined["10"];
echo $combined["20"];
echo isset($combined[30.0]) ? "has30" : "no30";
echo $combined[0] ?? "missing0";

__vybe_check(ob_get_clean(), "abno30missing0");
