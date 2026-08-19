<?php
// vybe-test: php/php_password_hashing_cryptography/test_php_hash_pbkdf2_key_derivation
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

$derived = hash_pbkdf2("sha256", "password", "salt", 1000, 32);
echo strlen($derived) === 64 ? "HEX_LEN_64" : "FAIL";


__vybe_check(ob_get_clean(), "FAIL");
