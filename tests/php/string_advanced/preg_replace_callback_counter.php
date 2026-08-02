<?php
// vybe-test: php/string_advanced/preg_replace_callback_counter
// origin: languages/php/tests/php/test_string_advanced.rs

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

$i = 0;
$result = preg_replace_callback('/\d+/', function($m) use (&$i) {
    $i++;
    return $m[0] * 2;
}, "a1 b2 c3");
echo $result;
echo "\n";
echo $i;
echo "\n";

__vybe_check(ob_get_clean(), "a2 b4 c6\n3");
