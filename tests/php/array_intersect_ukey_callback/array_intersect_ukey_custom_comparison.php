<?php
// vybe-test: php/array_intersect_ukey_callback/array_intersect_ukey_custom_comparison
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

$array1 = ['apple' => 1, 'banana' => 2];
$array2 = ['APPLE' => 3, 'Orange' => 4];

$result = array_intersect_ukey($array1, $array2, 'strcasecmp');
echo implode(',', array_keys($result));

__vybe_check(ob_get_clean(), "apple");
