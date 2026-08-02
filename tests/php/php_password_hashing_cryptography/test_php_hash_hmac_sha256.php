<?php
// vybe-test: php/php_password_hashing_cryptography/test_php_hash_hmac_sha256
// origin: languages/php/tests/php/test_php_password_hashing_cryptography.rs

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

$message = "payload_data";
$key = "secret_key_123";
$hmac = hash_hmac("sha256", $message, $key);
echo strlen($hmac) === 64 ? "HMAC_LENGTH_64" : "INVALID";

__vybe_check(ob_get_clean(), "HMAC_LENGTH_64");
