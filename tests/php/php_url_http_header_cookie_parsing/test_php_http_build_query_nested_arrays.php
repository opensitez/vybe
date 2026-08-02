<?php
// vybe-test: php/php_url_http_header_cookie_parsing/test_php_http_build_query_nested_arrays
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

$params = [
    "user" => ["name" => "Alice", "role" => "admin"],
    "filter" => "active"
];
echo http_build_query($params);

__vybe_check(ob_get_clean(), "user%5Bname%5D=Alice&user%5Brole%5D=admin&filter=active");
