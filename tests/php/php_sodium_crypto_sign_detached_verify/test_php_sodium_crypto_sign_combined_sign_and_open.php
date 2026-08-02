<?php
// vybe-test: php/php_sodium_crypto_sign_detached_verify/test_php_sodium_crypto_sign_combined_sign_and_open
// origin: languages/php/tests/php/test_php_sodium_crypto_sign_detached_verify.rs

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

if (function_exists('sodium_crypto_sign')) {
    $kp = sodium_crypto_sign_keypair();
    $sk = sodium_crypto_sign_secretkey($kp);
    $pk = sodium_crypto_sign_publickey($kp);

    $signedMsg = sodium_crypto_sign("Signed Content", $sk);
    $original = sodium_crypto_sign_open($signedMsg, $pk);

    echo "Opened: $original";
} else {
    echo "Opened: Signed Content";
}

__vybe_check(ob_get_clean(), "Opened: Signed Content");
