<?php
// vybe-test: php/array_intersect_ukey_callback/array_intersect_ukey_basic
// origin: languages/php/tests/php/test_array_intersect_ukey_callback.rs

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

$array1 = ['blue'  => 1, 'red'  => 2, 'green'  => 3, 'purple' => 4];
$array2 = ['green' => 5, 'blue' => 6, 'yellow' => 7, 'cyan'   => 8];

$result = array_intersect_ukey($array1, $array2, function ($key1, $key2) {
    if ($key1 == $key2) return 0;
    else if ($key1 > $key2) return 1;
    else return -1;
});

echo implode(',', array_keys($result));

__vybe_check(ob_get_clean(), "blue,green");
