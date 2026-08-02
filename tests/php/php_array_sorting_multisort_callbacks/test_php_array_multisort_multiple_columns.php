<?php
// vybe-test: php/php_array_sorting_multisort_callbacks/test_php_array_multisort_multiple_columns
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

$volume = [67, 86, 85, 98, 86, 67];
$edition = [2, 1, 6, 2, 1, 6];

array_multisort($volume, SORT_DESC, $edition, SORT_ASC);
echo "v0={$volume[0]} e0={$edition[0]} | v1={$volume[1]} e1={$edition[1]}";

__vybe_check(ob_get_clean(), "v0=98 e0=2 | v1=86 e1=1");
