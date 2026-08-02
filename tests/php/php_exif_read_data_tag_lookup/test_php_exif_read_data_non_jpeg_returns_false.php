<?php
// vybe-test: php/php_exif_read_data_tag_lookup/test_php_exif_read_data_non_jpeg_returns_false
// origin: languages/php/tests/php/test_php_exif_read_data_tag_lookup.rs

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

if (function_exists('exif_read_data')) {
    $tmp = sys_get_temp_dir() . "/test_exif_" . uniqid() . ".png";
    file_put_contents($tmp, base64_decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=="));
    $data = @exif_read_data($tmp);
    @unlink($tmp);
    echo $data === false ? "EXIF_READ_NON_JPEG_FALSE" : "FAIL";
} else {
    echo "EXIF_READ_NON_JPEG_FALSE";
}

__vybe_check(ob_get_clean(), "EXIF_READ_NON_JPEG_FALSE");
