<?php
// vybe-test: php/json_advanced/json_decode_to_assoc_array
// origin: languages/php/tests/php/test_json_advanced.rs

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

$a = json_decode('{"x":1,"y":2}', true);
echo $a['x'] + $a['y'];

__vybe_check(ob_get_clean(), "3");
