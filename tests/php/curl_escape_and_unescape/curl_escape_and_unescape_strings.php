<?php
// vybe-test: php/curl_escape_and_unescape/curl_escape_and_unescape_strings
// origin: languages/php/tests/php/test_curl_escape_and_unescape.rs

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

echo "curl_escape_and_unescape_strings_ok";

__vybe_check(ob_get_clean(), "curl_escape_and_unescape_strings_ok");
