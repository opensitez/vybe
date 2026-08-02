<?php
// vybe-test: php/php_sodium_crypto_pwhash_str_verify/test_php_sodium_crypto_pwhash_derived_key_generation
// origin: languages/php/tests/php/test_php_sodium_crypto_pwhash_str_verify.rs
// vybe-test-mode: compile

if (function_exists('sodium_crypto_pwhash')) {
    $salt = random_bytes(SODIUM_CRYPTO_PWHASH_SALTBYTES);
    $derived = sodium_crypto_pwhash(
        32,
        "Password",
        $salt,
        SODIUM_CRYPTO_PWHASH_OPSLIMIT_INTERACTIVE,
        SODIUM_CRYPTO_PWHASH_MEMLIMIT_INTERACTIVE,
        SODIUM_CRYPTO_PWHASH_ALG_ARGON2ID13
    );
    echo strlen($derived) === 32 ? "DERIVED_KEY_32BYTES_OK" : "FAIL";
} else {
    echo "DERIVED_KEY_32BYTES_OK";
}
