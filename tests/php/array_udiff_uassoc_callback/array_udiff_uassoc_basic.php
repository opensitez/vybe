<?php
// vybe-test: php/array_udiff_uassoc_callback/array_udiff_uassoc_basic
// origin: languages/php/tests/php/test_array_udiff_uassoc_callback.rs

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

$array1 = array("a" => "green", "b" => "brown", "c" => "blue", "red");
$array2 = array("a" => "GREEN", "B" => "brown", "yellow", "red");

$result = array_udiff_uassoc($array1, $array2, "strcasecmp", "strcasecmp");
ksort($result);
echo implode(',', array_keys($result));

__vybe_check(ob_get_clean(), "0,c");
