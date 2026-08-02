<?php
// vybe-test: php/hash_crypto_runtime/hash_md5_matches_md5_builtin
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

$h = hash('md5', 'hello');
echo (strlen($h) === 32 ? 'ok' : 'fail') . ($h === md5('hello') ? ':matches' : ':differs');

__vybe_check(ob_get_clean(), "ok:matches");
