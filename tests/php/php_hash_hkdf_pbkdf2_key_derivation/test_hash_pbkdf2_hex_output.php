<?php
// vybe-test: php/php_hash_hkdf_pbkdf2_key_derivation/test_hash_pbkdf2_hex_output
// origin: languages/php/tests/php/test_php_hash_hkdf_pbkdf2_key_derivation.rs

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

$hex = hash_pbkdf2('sha256', 'password', 'salt', 1000, 20, false);
echo strlen($hex), "\n";

__vybe_check(ob_get_clean(), "20");
