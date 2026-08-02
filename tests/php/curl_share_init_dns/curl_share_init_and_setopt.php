<?php
// vybe-test: php/curl_share_init_dns/curl_share_init_and_setopt
// origin: languages/php/tests/php/test_curl_share_init_dns.rs

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

$sh = curl_share_init();
$res1 = curl_share_setopt($sh, CURLSHOPT_SHARE, CURL_LOCK_DATA_COOKIE);
$res2 = curl_share_setopt($sh, CURLSHOPT_SHARE, CURL_LOCK_DATA_DNS);

echo ($res1 && $res2) ? "shared" : "failed";

$ch = curl_init("http://example.com");
curl_setopt($ch, CURLOPT_SHARE, $sh);
curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
curl_exec($ch);

curl_share_close($sh);

__vybe_check(ob_get_clean(), "shared");
