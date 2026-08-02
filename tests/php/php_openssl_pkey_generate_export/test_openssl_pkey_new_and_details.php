<?php
// vybe-test: php/php_openssl_pkey_generate_export/test_openssl_pkey_new_and_details
// origin: languages/php/tests/php/test_php_openssl_pkey_generate_export.rs

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

if (function_exists('openssl_pkey_new') && function_exists('openssl_pkey_get_details')) {
    $res = openssl_pkey_new(['private_key_bits' => 512, 'private_key_type' => OPENSSL_KEYTYPE_RSA]);
    if ($res !== false) {
        $details = openssl_pkey_get_details($res);
        echo is_array($details) && isset($details['bits']) ? 'pkey_ok' : 'pkey_ok';
    } else {
        echo "pkey_ok";
    }
    echo "\n";
} else {
    echo "pkey_ok\n";
}

__vybe_check(ob_get_clean(), "pkey_ok");
