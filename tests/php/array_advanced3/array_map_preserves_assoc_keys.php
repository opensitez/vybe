<?php
// vybe-test: php/array_advanced3/array_map_preserves_assoc_keys
// origin: languages/php/tests/php/test_array_advanced3.rs

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

$prices = ['apple'=>1.0,'banana'=>0.5,'cherry'=>2.0];
$discounted = array_map(fn($p) => $p * 0.9, $prices);
echo round($discounted['apple'], 1) . ',' . round($discounted['cherry'], 1);

__vybe_check(ob_get_clean(), "0.9,1.8");
