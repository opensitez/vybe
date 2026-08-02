<?php
// vybe-test: php/array_creation/array_rand_multiple_ordered_without_shuffle
// origin: languages/php/tests/php/test_array_creation.rs

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

$r = array_rand(['a' => 1, 'b' => 2, 'c' => 3], 2);
sort($r);
echo $r[0] . $r[1];

__vybe_check(ob_get_clean(), "ab");
