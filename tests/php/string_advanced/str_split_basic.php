<?php
// vybe-test: php/string_advanced/str_split_basic
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

$chars = str_split("hello");
echo implode(",", $chars);
echo "\n";
$chunks = str_split("abcdefgh", 3);
echo implode("|", $chunks);
echo "\n";

__vybe_check(ob_get_clean(), "h,e,l,l,o\nabc|def|gh");
