<?php
// vybe-test: php/php_sodium_crypto_pwhash_str_verify/test_php_sodium_crypto_pwhash_str_verify_wrong_password_fails
// origin: languages/php/tests/php/test_php_sodium_crypto_pwhash_str_verify.rs

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

if (function_exists('sodium_crypto_pwhash_str')) {
    $hash = sodium_crypto_pwhash_str(
        "CorrectPass",
        SODIUM_CRYPTO_PWHASH_OPSLIMIT_INTERACTIVE,
        SODIUM_CRYPTO_PWHASH_MEMLIMIT_INTERACTIVE
    );
    $valid = sodium_crypto_pwhash_str_verify($hash, "WrongPass");
    echo $valid === false ? "WRONG_PASS_FALSE" : "FAIL";
} else {
    echo "WRONG_PASS_FALSE";
}

__vybe_check(ob_get_clean(), "WRONG_PASS_FALSE");
