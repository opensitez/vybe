<?php
// vybe-test: php/php_curl_http_client_requests/test_php_curl_setopt_array_batch_configuration
// origin: languages/php/tests/php/test_php_curl_http_client_requests.rs

function __vybe_check($got, $want) {
    $got = str_replace("\r", "", $got);
    $got = rtrim($got, "\n");
    if ($got !== $want) {
        echo "FAIL: want [" . $want . "] got [" . $got . "]\n";
        throw new Exception("assertion failed");
    }
    echo $got;
    if ($got !== "") {
        echo "\n";
    }
}

ob_start();

echo "test_php_curl_setopt_array_batch_configuration_ok";

__vybe_check(ob_get_clean(), "test_php_curl_setopt_array_batch_configuration_ok");
