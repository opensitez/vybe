<?php
// vybe-test: php/php_array_diff_intersect_associations/test_php_array_diff_assoc_keys_and_values
// origin: languages/php/tests/php/test_php_array_diff_intersect_associations.rs

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

$a1 = ["a" => "green", "b" => "brown", "c" => "blue", "red"];
$a2 = ["a" => "green", "yellow", "red"];
$diff = array_diff_assoc($a1, $a2);
echo implode(",", array_keys($diff));

__vybe_check(ob_get_clean(), "b,c,0");
