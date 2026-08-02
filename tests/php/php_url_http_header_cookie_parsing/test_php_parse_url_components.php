<?php
// vybe-test: php/php_url_http_header_cookie_parsing/test_php_parse_url_components
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

$url = "https://user:pass@example.com:8080/path/to/script.php?param=val#section";
$parsed = parse_url($url);

echo $parsed["scheme"] . " | " . $parsed["host"] . " | " . $parsed["port"] . " | " . $parsed["path"];

__vybe_check(ob_get_clean(), "https | example.com | 8080 | /path/to/script.php");
