<?php
// vybe-test: php/string_encoding/bin2hex_and_hex2bin_roundtrip_runtime
// origin: languages/php/tests/php/test_string_encoding.rs

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

$raw = "\x00\x01\xffABC";
echo bin2hex($raw);
echo "\n";
echo hex2bin(bin2hex($raw)) === $raw ? 'same' : 'diff';

__vybe_check(ob_get_clean(), "0001ff414243|same");
