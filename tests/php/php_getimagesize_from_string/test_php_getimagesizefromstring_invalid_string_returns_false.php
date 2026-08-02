<?php
// vybe-test: php/php_getimagesize_from_string/test_php_getimagesizefromstring_invalid_string_returns_false
// origin: languages/php/tests/php/test_php_getimagesize_from_string.rs

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

if (function_exists('getimagesizefromstring')) {
    $info = @getimagesizefromstring("not an image binary payload");
    echo $info === false ? "INVALID_IMAGE_FALSE" : "FAIL";
} else {
    echo "INVALID_IMAGE_FALSE";
}

__vybe_check(ob_get_clean(), "INVALID_IMAGE_FALSE");
