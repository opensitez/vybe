<?php
// vybe-test: php/array_advanced/array_map_null_callback_zip
// origin: languages/php/tests/php/test_array_advanced.rs

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

$a = [1, 2, 3];
$b = ["a", "b", "c"];
$zipped = array_map(null, $a, $b);
foreach ($zipped as $pair) {
    echo $pair[0] . $pair[1];
}

__vybe_check(ob_get_clean(), "1a2b3c");
