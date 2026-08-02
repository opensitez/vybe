<?php
// vybe-test: php/type_juggling_runtime/settype_invalid_to_invalid
// origin: languages/php/tests/php/test_type_juggling_runtime.rs

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

$value = 'abc';
settype($value, 'integer');
echo $value . '|';
$value = 1.9;
settype($value, 'string');
echo $value;

__vybe_check(ob_get_clean(), "0|1.9");
