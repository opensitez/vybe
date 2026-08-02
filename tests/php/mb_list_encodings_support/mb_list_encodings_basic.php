<?php
// vybe-test: php/mb_list_encodings_support/mb_list_encodings_basic
// origin: languages/php/tests/php/test_mb_list_encodings_support.rs

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

$list = mb_list_encodings();
echo in_array("UTF-8", $list) ? "found" : "missing";

__vybe_check(ob_get_clean(), "found");
