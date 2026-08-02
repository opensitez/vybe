<?php
// vybe-test: php/php_hash_hkdf_pbkdf2_key_derivation/test_hash_hkdf_length_and_hex
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

$derived = hash_hkdf('sha256', 'secret_ikm', 32, 'app_info', 'salt123');
echo strlen($derived) . ':' . bin2hex(substr($derived, 0, 4)), "\n";

__vybe_check(ob_get_clean(), "32:a816bd4d");
