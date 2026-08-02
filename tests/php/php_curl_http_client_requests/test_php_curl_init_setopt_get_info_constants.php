<?php
// vybe-test: php/php_curl_http_client_requests/test_php_curl_init_setopt_get_info_constants
// origin: languages/php/tests/php/test_php_curl_http_client_requests.rs

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

$ch = curl_init("https://example.com/api/test");
curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
curl_setopt($ch, CURLOPT_TIMEOUT, 10);
curl_setopt($ch, CURLOPT_HTTPHEADER, ["Accept: application/json"]);

$info = curl_getinfo($ch);
curl_close($ch);

echo "URL={$info['url']} Opts_Set=1";

__vybe_check(ob_get_clean(), "URL=https://example.com/api/test Opts_Set=1");
