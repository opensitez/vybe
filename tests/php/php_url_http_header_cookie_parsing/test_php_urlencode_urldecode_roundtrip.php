<?php
// vybe-test: php/php_url_http_header_cookie_parsing/test_php_urlencode_urldecode_roundtrip
// origin: languages/php/tests/php/test_php_url_http_header_cookie_parsing.rs

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

$raw = "Parameter & Value with spaces / slashes";
$encoded = urlencode($raw);
$decoded = urldecode($encoded);
echo ($decoded === $raw ? "ROUNDTRIP_OK" : "ROUNDTRIP_FAIL");

__vybe_check(ob_get_clean(), "ROUNDTRIP_OK");
