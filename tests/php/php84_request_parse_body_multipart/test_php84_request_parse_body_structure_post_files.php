<?php
// vybe-test: php/php84_request_parse_body_multipart/test_php84_request_parse_body_structure_post_files
// origin: languages/php/tests/php/test_php84_request_parse_body_multipart.rs

function __vybe_check($got, $want) {
    $got = str_replace("\r", "", $got);
    $got = rtrim($got, "\n");
    if ($got !== $want) {
        echo "FAIL: want [" . $want . "] got [" . $got . "]\n";
        throw new Exception("assertion failed");
    }
    echo $got;
    if ($got !== "") {
        echo "\n";
    }
}

ob_start();

echo "test_php84_request_parse_body_structure_post_files_ok";

__vybe_check(ob_get_clean(), "test_php84_request_parse_body_structure_post_files_ok");
