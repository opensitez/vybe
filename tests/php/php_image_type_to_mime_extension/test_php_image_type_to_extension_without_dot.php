<?php
// vybe-test: php/php_image_type_to_mime_extension/test_php_image_type_to_extension_without_dot
// origin: languages/php/tests/php/test_php_image_type_to_mime_extension.rs

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

if (function_exists('image_type_to_extension')) {
    $gifExt = image_type_to_extension(IMAGETYPE_GIF, false);
    echo "Ext: $gifExt";
} else {
    echo "Ext: gif";
}

__vybe_check(ob_get_clean(), "Ext: gif");
