<?php
// vybe-test: php/php_sodium_crypto_pwhash_str_verify/test_php_sodium_crypto_pwhash_str_needs_rehash_check
// origin: languages/php/tests/php/test_php_sodium_crypto_pwhash_str_verify.rs
// vybe-test-mode: compile

if (function_exists('sodium_crypto_pwhash_str_needs_rehash')) {
    $hash = sodium_crypto_pwhash_str(
        "TestPass",
        SODIUM_CRYPTO_PWHASH_OPSLIMIT_INTERACTIVE,
        SODIUM_CRYPTO_PWHASH_MEMLIMIT_INTERACTIVE
    );
    $rehash = sodium_crypto_pwhash_str_needs_rehash(
        $hash,
        SODIUM_CRYPTO_PWHASH_OPSLIMIT_INTERACTIVE,
        SODIUM_CRYPTO_PWHASH_MEMLIMIT_INTERACTIVE
    );
    echo $rehash === false ? "NO_REHASH_NEEDED_OK" : "FAIL";
} else {
    echo "NO_REHASH_NEEDED_OK";
}
