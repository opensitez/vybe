<?php
// vybe-test: php/php_string_hex_base64_hashing/test_php_base64_encode_and_decode_roundtrip
// origin: languages/php/tests/php/test_php_string_hex_base64_hashing.rs

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

$original = "Hello World! Binary \x00\x01\x02";
$encoded = base64_encode($original);
$decoded = base64_decode($encoded);

echo ($decoded === $original ? "BASE64_ROUNDTRIP_OK" : "FAIL");

__vybe_check(ob_get_clean(), "BASE64_ROUNDTRIP_OK");
