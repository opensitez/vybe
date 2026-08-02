<?php
// vybe-test: php/json_errors/json_encode_hex_tag_escapes_angle_brackets
// origin: languages/php/tests/php/test_json_errors.rs

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

$out = json_encode('<tag>', JSON_THROW_ON_ERROR | JSON_HEX_TAG);
echo str_contains($out, '\\u003C') ? 'hex' : 'plain';

__vybe_check(ob_get_clean(), "hex");
