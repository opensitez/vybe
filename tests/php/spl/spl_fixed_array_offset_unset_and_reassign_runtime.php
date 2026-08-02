<?php
// vybe-test: php/spl/spl_fixed_array_offset_unset_and_reassign_runtime
// origin: languages/php/tests/php/test_spl.rs

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

$arr = new SplFixedArray(3);
$arr[0] = 'first';
$arr[1] = 'second';
$arr[2] = 'third';
unset($arr[1]);
$arr[1] = 'rebound';
echo $arr->count();
echo '|';
echo $arr[1];

__vybe_check(ob_get_clean(), "3|rebound");
