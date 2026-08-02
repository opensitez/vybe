<?php
// vybe-test: php/php_web_http_response_code_getter_setter/test_http_response_code_default_200
// origin: languages/php/tests/php/test_php_web_http_response_code_getter_setter.rs

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

if (function_exists('http_response_code')) {
    $code = http_response_code();
    echo ($code === 200 || $code === false) ? 'default_code_ok' : 'err', "\n";
} else {
    echo "default_code_ok\n";
}

__vybe_check(ob_get_clean(), "default_code_ok");
