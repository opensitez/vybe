<?php
// vybe-test: php/php_json_encode_decode_options/test_php_json_decode_bigint_as_string
// origin: languages/php/tests/php/test_php_json_encode_decode_options.rs

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

$json = '{"big_int": 9223372036854775807}';
$data = json_decode($json, true, flags: JSON_BIGINT_AS_STRING);
echo gettype($data["big_int"]) . "=" . $data["big_int"];

__vybe_check(ob_get_clean(), "integer=9223372036854775807");
