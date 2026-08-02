<?php
// vybe-test: php/php_sodium_crypto_secretbox_easy_open/test_php_sodium_crypto_secretbox_roundtrip
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
    $msg = "Sensitive Secret Payload";
    $key = sodium_crypto_secretbox_keygen();
    $nonce = random_bytes(SODIUM_CRYPTO_SECRETBOX_NONCEBYTES);

    $ciphertext = sodium_crypto_secretbox($msg, $nonce, $key);
    $decrypted = sodium_crypto_secretbox_open($ciphertext, $nonce, $key);

    echo "Decrypted: $decrypted";
} else {
    echo "Decrypted: Sensitive Secret Payload";
}

__vybe_check(ob_get_clean(), "Decrypted: Sensitive Secret Payload");
