<?php
// vybe-test: php/php_web_parse_url_invalid_url_handling/test_parse_url_scheme_relative
// origin: languages/php/tests/php/test_php_web_parse_url_invalid_url_handling.rs

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

$res = parse_url('//cdn.example.com/app.js');
echo $res['host'] . ':' . $res['path'], "\n";

__vybe_check(ob_get_clean(), "cdn.example.com:/app.js");
