<?php
// vybe-test: php/php_image_type_to_mime_extension/test_php_image_type_to_mime_type_bmp
// origin: languages/php/tests/php/test_php_image_type_to_mime_extension.rs
// vybe-test-mode: compile

if (function_exists('image_type_to_mime_type') && defined('IMAGETYPE_BMP')) {
    $mime = image_type_to_mime_type(IMAGETYPE_BMP);
    echo $mime === "image/bmp" || $mime === "image/x-ms-bmp" ? "BMP_MIME_OK" : "FAIL";
} else {
    echo "BMP_MIME_OK";
}
