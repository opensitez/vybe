<?php
// vybe-test: php/array_functions_extra/array_push_and_unshift_return_counts
// origin: languages/php/tests/php/test_array_functions_extra.rs

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

$a = [1];
$pushed = array_push($a, 2, 3);
$unshifted = array_unshift($a, 0);
echo $pushed;
echo '|';
echo $unshifted;
echo '|';
echo implode(',', $a);

__vybe_check(ob_get_clean(), "3|4|0,1,2,3");
