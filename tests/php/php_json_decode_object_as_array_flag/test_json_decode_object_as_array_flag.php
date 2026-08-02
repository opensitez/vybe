<?php
// vybe-test: php/php_json_decode_object_as_array_flag/test_json_decode_object_as_array_flag
// origin: languages/php/tests/php/test_php_json_decode_object_as_array_flag.rs

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

$json = '{"name":"Alice","items":[1,2,3]}';
$decoded = json_decode($json, false, 512, JSON_OBJECT_AS_ARRAY);
echo is_array($decoded) && $decoded['name'] === 'Alice' ? 'array_decoded' : 'err', "\n";

__vybe_check(ob_get_clean(), "array_decoded");
