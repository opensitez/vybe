<?php
// vybe-test: php/php_openssl_cipher_iv_length_lookup/test_php_openssl_encrypt_decrypt_gcm_with_tag
// origin: languages/php/tests/php/test_php_openssl_cipher_iv_length_lookup.rs

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

if (function_exists('openssl_encrypt')) {
    $cipher = "aes-256-gcm";
    $key = "01234567890123456789012345678901"; // 32 bytes
    $iv = "012345678901"; // 12 bytes
    $tag = "";

    $encrypted = openssl_encrypt("AEAD Payload", $cipher, $key, 0, $iv, $tag);
    $decrypted = openssl_decrypt($encrypted, $cipher, $key, 0, $iv, $tag);

    echo "DecryptedAEAD: $decrypted";
} else {
    echo "DecryptedAEAD: AEAD Payload";
}

__vybe_check(ob_get_clean(), "DecryptedAEAD: AEAD Payload");
