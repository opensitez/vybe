<?php
// vybe-test: php/random_crypto/bin2hex_random_bytes_is_hex_string
// origin: languages/php/tests/php/test_random_crypto.rs

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

$h = bin2hex(random_bytes(3));
echo ctype_xdigit($h) && strlen($h) === 6 ? 'hex' : 'no';

__vybe_check(ob_get_clean(), "hex");
