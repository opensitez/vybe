<?php
// vybe-test: php/hash_crypto_runtime/crc32_is_int_and_deterministic
// origin: languages/php/tests/php/test_hash_crypto_runtime.rs

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

$checksum = crc32('hello world');
echo (is_int($checksum) ? 'int' : 'not int') . (crc32('hello world') === $checksum ? ':deterministic' : ':varies');

__vybe_check(ob_get_clean(), "int:deterministic");
