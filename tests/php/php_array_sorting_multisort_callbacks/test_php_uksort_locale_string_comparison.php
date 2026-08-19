<?php
// vybe-test: php/php_array_sorting_multisort_callbacks/test_php_uksort_locale_string_comparison
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

$data = ["éclair" => 1, "apple" => 2, "Éclair" => 3, "banana" => 4];
uksort($data, fn($a, $b) => strcmp($a, $b));
echo implode("|", array_keys($data));

__vybe_check(ob_get_clean(), "apple|banana|Éclair|éclair");
