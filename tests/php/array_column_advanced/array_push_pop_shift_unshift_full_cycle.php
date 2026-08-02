<?php
// vybe-test: php/array_column_advanced/array_push_pop_shift_unshift_full_cycle
// origin: languages/php/tests/php/test_array_column_advanced.rs

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

$q = [2];
array_unshift($q, 1);
array_push($q, 3);
$x = array_shift($q);
$y = array_pop($q);
echo $x . $y . '|' . implode(',', $q);

__vybe_check(ob_get_clean(), "13|2");
