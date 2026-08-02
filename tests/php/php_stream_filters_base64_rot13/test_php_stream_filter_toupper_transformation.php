<?php
// vybe-test: php/php_stream_filters_base64_rot13/test_php_stream_filter_toupper_transformation
// origin: languages/php/tests/php/test_php_stream_filters_base64_rot13.rs

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

$stream = fopen("php://memory", "r+");
stream_filter_append($stream, "string.toupper");

fwrite($stream, "lowercase text");
rewind($stream);

$upper = stream_get_contents($stream);
fclose($stream);

echo $upper;

__vybe_check(ob_get_clean(), "LOWERCASE TEXT");
