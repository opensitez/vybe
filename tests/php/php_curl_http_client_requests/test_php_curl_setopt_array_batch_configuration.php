<?php
// vybe-test: php/php_curl_http_client_requests/test_php_curl_setopt_array_batch_configuration
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

$ch = curl_init();
$options = [
    CURLOPT_URL => "https://api.github.com/users",
    CURLOPT_USERAGENT => "Vybe-Client/1.0",
    CURLOPT_RETURNTRANSFER => true,
    CURLOPT_HEADER => false,
];

curl_setopt_array($ch, $options);
$url = curl_getinfo($ch, CURLINFO_EFFECTIVE_URL);
curl_close($ch);

echo "Effective URL: $url";

__vybe_check(ob_get_clean(), "Effective URL: https://api.github.com/users");
