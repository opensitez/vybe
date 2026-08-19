<?php
// vybe-test: php/string_encoding/json_encode_escape_avoids_control
// origin: languages/php/tests/php/test_string_encoding.rs

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

$value = ["path" => "café", "n" => 2];
$json = json_encode($value, JSON_UNESCAPED_UNICODE);
echo $json;
echo "\n";
echo json_decode($json, true)['path'];

__vybe_check(ob_get_clean(), "{\"path\":\"café\",\"n\":2}\ncafé");
