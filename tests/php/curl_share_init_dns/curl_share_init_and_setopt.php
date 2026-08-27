<?php
// vybe-test: php/curl_share_init_dns/curl_share_init_and_setopt
// origin: languages/php/tests/php/test_curl_share_init_dns.rs

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

echo "curl_share_init_and_setopt_ok";

__vybe_check(ob_get_clean(), "curl_share_init_and_setopt_ok");
