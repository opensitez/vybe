<?php
// vybe-test: php/array_callbacks/array_map_with_null_callback_zip
// origin: languages/php/tests/php/test_array_callbacks.rs

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

$letters = ['a', 'b', 'c'];
$numbers = [1, 2, 3];
$zipped = array_map(null, $letters, $numbers);
echo json_encode($zipped);

__vybe_check(ob_get_clean(), "[[\"a\",1],[\"b\",2],[\"c\",3]]");
