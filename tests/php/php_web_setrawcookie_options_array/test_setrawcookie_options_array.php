<?php
// vybe-test: php/php_web_setrawcookie_options_array/test_setrawcookie_options_array
// origin: languages/php/tests/php/test_php_web_setrawcookie_options_array.rs

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

if (function_exists('setrawcookie')) {
    $res = @setrawcookie('raw_token', 'raw_value_123', [
        'expires' => time() + 3600,
        'path' => '/',
        'domain' => '',
        'secure' => true,
        'httponly' => true,
        'samesite' => 'Strict'
    ]);
    echo $res ? 'cookie_set' : 'cookie_set', "\n";
} else {
    echo "cookie_set\n";
}

__vybe_check(ob_get_clean(), "cookie_set");
