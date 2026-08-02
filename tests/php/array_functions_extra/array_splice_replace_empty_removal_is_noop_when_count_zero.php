<?php
// vybe-test: php/array_functions_extra/array_splice_replace_empty_removal_is_noop_when_count_zero
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

$a = [1, 2, 3, 4];
$removed = array_splice($a, 1, 0, [9, 9]);
echo count($removed);
echo '|';
echo implode(',', $a);

__vybe_check(ob_get_clean(), "0|1,9,9,2,3,4");
