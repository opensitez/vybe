<?php
// vybe-test: php/str_getcsv_enclosures/str_getcsv_custom_delimiter_enclosure
// origin: languages/php/tests/php/test_str_getcsv_enclosures.rs

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

$str = "123;'hello;world';456";
$arr = str_getcsv($str, ';', "'");
echo count($arr) . "|" . $arr[1];

__vybe_check(ob_get_clean(), "3|hello;world");
