<?php
// vybe-test: php/array_builtins_extended/array_filter_use_value_and_key_modes
// origin: languages/php/tests/php/test_array_builtins_extended.rs

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

$data = ["a" => 1, "b" => 0, "c" => 3, "d" => "0"];
$vals = array_filter($data, fn($v) => $v, 0);
$keys = array_filter($data, fn($k) => $k === "a" || $k === "d", ARRAY_FILTER_USE_KEY);
echo count($vals) . "|" . count($keys);

__vybe_check(ob_get_clean(), "2|2");
