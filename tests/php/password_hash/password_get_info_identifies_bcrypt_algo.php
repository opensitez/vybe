<?php
// vybe-test: php/password_hash/password_get_info_identifies_bcrypt_algo
// origin: languages/php/tests/php/test_password_hash.rs

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

$h = password_hash('x', PASSWORD_BCRYPT);
$info = password_get_info($h);
echo ($info['algoName'] ?? '') === 'bcrypt' ? 'bcrypt' : 'other';

__vybe_check(ob_get_clean(), "bcrypt");
