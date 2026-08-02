<?php
// vybe-test: php/curl_multi_exec_loop/curl_multi_exec_loop_execution
// origin: languages/php/tests/php/test_curl_multi_exec_loop.rs

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
$ch = curl_init("http://example.com");
curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
curl_multi_add_handle($mh, $ch);

$active = null;
do {
    $status = curl_multi_exec($mh, $active);
    if ($active) {
        curl_multi_select($mh);
    }
} while ($active && $status == CURLM_OK);

$content = curl_multi_getcontent($ch);
echo is_string($content) && strlen($content) > 0 ? "success" : "failed";

curl_multi_remove_handle($mh, $ch);
curl_multi_close($mh);

__vybe_check(ob_get_clean(), "success");
