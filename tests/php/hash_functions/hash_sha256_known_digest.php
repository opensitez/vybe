<?php
// vybe-test: php/hash_functions/hash_sha256_known_digest
// origin: languages/php/tests/php/test_hash_functions.rs

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

echo hash('sha256', 'abc');

__vybe_check(ob_get_clean(), "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad");
