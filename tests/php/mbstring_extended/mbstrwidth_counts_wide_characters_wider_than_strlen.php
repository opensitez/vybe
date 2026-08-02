<?php
// vybe-test: php/mbstring_extended/mbstrwidth_counts_wide_characters_wider_than_strlen
// origin: languages/php/tests/php/test_mbstring_extended.rs

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

$s = '日本';
echo mb_strwidth($s) . ':' . strlen($s);

__vybe_check(ob_get_clean(), "4:6");
