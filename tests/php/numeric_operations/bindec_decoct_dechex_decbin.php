<?php
// vybe-test: php/numeric_operations/bindec_decoct_dechex_decbin
// origin: languages/php/tests/php/test_numeric_operations.rs

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

echo bindec('1010') . "\n";
echo decbin(10) . "\n";
echo octdec('77') . "\n";
echo decoct(63) . "\n";
echo hexdec('ff') . "\n";
echo dechex(255) . "\n";

__vybe_check(ob_get_clean(), "10\n1010\n63\n77\n255\nff");
