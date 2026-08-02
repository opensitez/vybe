<?php
// vybe-test: php/php_type_juggling_coercion_strictness/test_php_explicit_type_casting_syntax
// origin: languages/php/tests/php/test_php_type_juggling_coercion_strictness.rs

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

$val = "123.45";
$asInt = (int)$val;
$asFloat = (float)$val;
$asBool = (bool)$val;
$asArr = (array)$val;
$asObj = (object)$val;

echo "$asInt | $asFloat | " . ($asBool ? "1" : "0") . " | " . $asArr[0] . " | " . $asObj->scalar;

__vybe_check(ob_get_clean(), "123 | 123.45 | 1 | 123.45 | 123.45");
