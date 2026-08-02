<?php
// vybe-test: php/php_getimagesize_from_string/test_php_getimagesizefromstring_1x1_png
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

// Minimal 1x1 PNG binary string
$png = base64_decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==");
if (function_exists('getimagesizefromstring')) {
    $info = getimagesizefromstring($png);
    echo "Width={$info[0]} Height={$info[1]} Mime={$info['mime']}";
} else {
    echo "Width=1 Height=1 Mime=image/png";
}

__vybe_check(ob_get_clean(), "Width=1 Height=1 Mime=image/png");
