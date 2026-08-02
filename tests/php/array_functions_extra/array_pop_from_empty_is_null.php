<?php
// vybe-test: php/array_functions_extra/array_pop_from_empty_is_null
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

$a = [];
$v = array_pop($a);
echo is_null($v) ? 'null' : 'notnull';
echo '|';
echo is_array($a) ? 'is_array' : 'no';
echo '|';
echo count($a);

__vybe_check(ob_get_clean(), "null|is_array|0");
