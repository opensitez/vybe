<?php
// vybe-test: php/php_password_hashing_cryptography/test_php_password_needs_rehash_options
// origin: languages/php/tests/php/test_php_password_hashing_cryptography.rs

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

$hash = password_hash("secret", PASSWORD_BCRYPT, ["cost" => 4]);
$needs = password_needs_rehash($hash, PASSWORD_BCRYPT, ["cost" => 10]);
echo $needs ? "NEEDS_REHASH" : "OK";

__vybe_check(ob_get_clean(), "NEEDS_REHASH");
