<?php
// vybe-test: php/array_advanced/array_search_strict_mode
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

$a = ["10", "20", "30"];
$k = array_search(20, $a, true);
echo ($k === false) ? "not found" : $k;
$k2 = array_search("20", $a, true);
echo $k2;

__vybe_check(ob_get_clean(), "not found1");
