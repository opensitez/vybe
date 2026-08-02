<?php
// vybe-test: php/curl_escape_and_unescape/curl_escape_and_unescape_strings
// origin: languages/php/tests/php/test_curl_escape_and_unescape.rs

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
$escaped = curl_escape($ch, "hello world = +");
$unescaped = curl_unescape($ch, $escaped);

echo $escaped . "|" . $unescaped;
curl_close($ch);

__vybe_check(ob_get_clean(), "hello%20world%20%3D%20%2B|hello world = +");
