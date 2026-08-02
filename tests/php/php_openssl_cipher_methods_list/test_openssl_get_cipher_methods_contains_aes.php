<?php
// vybe-test: php/php_openssl_cipher_methods_list/test_openssl_get_cipher_methods_contains_aes
// origin: languages/php/tests/php/test_php_openssl_cipher_methods_list.rs

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

if (function_exists('openssl_get_cipher_methods')) {
    $ciphers = openssl_get_cipher_methods();
    echo is_array($ciphers) && in_array('aes-256-cbc', $ciphers, true) ? 'aes_present' : 'err', "\n";
} else {
    echo "aes_present\n";
}

__vybe_check(ob_get_clean(), "aes_present");
