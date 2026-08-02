<?php
// vybe-test: php/curl_multi_init_and_add/curl_multi_init_and_add_handle
// origin: languages/php/tests/php/test_curl_multi_init_and_add.rs

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

$mh = curl_multi_init();
$ch1 = curl_init("http://example.com");
$ch2 = curl_init("http://example.org");

$code1 = curl_multi_add_handle($mh, $ch1);
$code2 = curl_multi_add_handle($mh, $ch2);

// 0 is CURLM_OK
echo ($code1 === 0 && $code2 === 0) ? "added" : "failed";

curl_multi_remove_handle($mh, $ch1);
curl_multi_remove_handle($mh, $ch2);
curl_multi_close($mh);

__vybe_check(ob_get_clean(), "added");
