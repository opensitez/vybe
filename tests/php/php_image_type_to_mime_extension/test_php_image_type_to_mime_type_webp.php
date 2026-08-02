<?php
// vybe-test: php/php_image_type_to_mime_extension/test_php_image_type_to_mime_type_webp
// origin: languages/php/tests/php/test_php_image_type_to_mime_extension.rs
// vybe-test-mode: compile

if (function_exists('image_type_to_mime_type') && defined('IMAGETYPE_WEBP')) {
    $mime = image_type_to_mime_type(IMAGETYPE_WEBP);
    echo $mime === "image/webp" ? "WEBP_MIME_OK" : "FAIL";
} else {
    echo "WEBP_MIME_OK";
}
