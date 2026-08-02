<?php
// vybe-test: php/array_advanced/array_key_last_with_all_reference_keys
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

$data = [];
$data[] = 10;
$data["1"] = 20;
$data["alpha"] = 30;
echo array_key_first($data);
echo "|";
echo array_key_last($data);
echo "|";
echo count($data);

__vybe_check(ob_get_clean(), "0|alpha|3");
