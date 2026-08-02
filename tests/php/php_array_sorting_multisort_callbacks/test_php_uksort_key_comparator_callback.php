<?php
// vybe-test: php/php_array_sorting_multisort_callbacks/test_php_uksort_key_comparator_callback
// origin: languages/php/tests/php/test_php_array_sorting_multisort_callbacks.rs

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

$data = ["2024-05-01" => "a", "2024-01-01" => "b", "2024-12-01" => "c"];
uksort($data, fn($k1, $k2) => strtotime($k1) <=> strtotime($k2));
echo implode(",", array_keys($data));

__vybe_check(ob_get_clean(), "2024-01-01,2024-05-01,2024-12-01");
