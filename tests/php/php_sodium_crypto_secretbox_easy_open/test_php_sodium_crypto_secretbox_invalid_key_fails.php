<?php
// vybe-test: php/php_sodium_crypto_secretbox_easy_open/test_php_sodium_crypto_secretbox_invalid_key_fails
// origin: languages/php/tests/php/test_php_sodium_crypto_secretbox_easy_open.rs

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

if (function_exists('sodium_crypto_secretbox')) {
    $msg = "Payload";
    $k1 = sodium_crypto_secretbox_keygen();
    $k2 = sodium_crypto_secretbox_keygen();
    $nonce = random_bytes(SODIUM_CRYPTO_SECRETBOX_NONCEBYTES);

    $ciphertext = sodium_crypto_secretbox($msg, $nonce, $k1);
    $res = sodium_crypto_secretbox_open($ciphertext, $nonce, $k2);

    echo $res === false ? "DECRYPT_FAILED" : "FAIL";
} else {
    echo "DECRYPT_FAILED";
}

__vybe_check(ob_get_clean(), "DECRYPT_FAILED");
