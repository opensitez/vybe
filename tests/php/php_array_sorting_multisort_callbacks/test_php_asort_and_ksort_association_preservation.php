<?php
// vybe-test: php/php_array_sorting_multisort_callbacks/test_php_asort_and_ksort_association_preservation
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

$fruit = ["d" => "lemon", "a" => "orange", "b" => "banana", "c" => "apple"];
asort($fruit);
echo implode(",", array_keys($fruit)) . " | " . implode(",", $fruit);

__vybe_check(ob_get_clean(), "c,b,d,a | apple,banana,lemon,orange");
