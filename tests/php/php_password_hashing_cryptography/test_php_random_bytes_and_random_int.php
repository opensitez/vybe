<?php
// vybe-test: php/php_password_hashing_cryptography/test_php_random_bytes_and_random_int
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

$bytes = random_bytes(16);
$randNum = random_int(1, 100);

echo (strlen($bytes) === 16 && $randNum >= 1 && $randNum <= 100) ? "CSPRNG_OK" : "FAIL";

__vybe_check(ob_get_clean(), "CSPRNG_OK");
