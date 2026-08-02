<?php
// vybe-test: php/php_web_http_build_query_enc_type/test_http_build_query_rfc3986_encoding
// origin: languages/php/tests/php/test_php_web_http_build_query_enc_type.rs

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

$data = ['name' => 'John Doe', 'symbol' => 'foo+bar'];
echo http_build_query($data, '', '&', PHP_QUERY_RFC3986), "\n";

__vybe_check(ob_get_clean(), "name=John%20Doe&symbol=foo%2Bbar");
