<?php
// vybe-test: php/url_http_runtime/http_response_code_returns_int_or_false
// origin: languages/php/tests/php/test_url_http_runtime.rs

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

$code = http_response_code();
echo is_int($code) || $code === false ? 'ok' : 'fail';

__vybe_check(ob_get_clean(), "ok");
