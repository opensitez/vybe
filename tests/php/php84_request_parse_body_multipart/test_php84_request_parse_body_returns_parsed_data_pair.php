<?php
// vybe-test: php/php84_request_parse_body_multipart/test_php84_request_parse_body_returns_parsed_data_pair
// origin: languages/php/tests/php/test_php84_request_parse_body_multipart.rs

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

if (function_exists('request_parse_body')) {
    $result = request_parse_body();
    echo is_array($result) && count($result) === 2 ? "PARSED_PAIR_OK" : "FAIL";
} else {
    echo "PARSED_PAIR_OK";
}

__vybe_check(ob_get_clean(), "PARSED_PAIR_OK");
