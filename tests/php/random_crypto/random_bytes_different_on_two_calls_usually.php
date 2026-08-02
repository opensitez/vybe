<?php
// vybe-test: php/random_crypto/random_bytes_different_on_two_calls_usually
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

$a = bin2hex(random_bytes(4));
$b = bin2hex(random_bytes(4));
echo strlen($a) === 8 && strlen($b) === 8 ? 'ok' : 'bad';

__vybe_check(ob_get_clean(), "ok");
