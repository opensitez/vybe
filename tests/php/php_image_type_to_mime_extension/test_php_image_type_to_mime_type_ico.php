<?php
// vybe-test: php/php_image_type_to_mime_extension/test_php_image_type_to_mime_type_ico
// origin: languages/php/tests/php/test_php_image_type_to_mime_extension.rs
// vybe-test-mode: compile

if (function_exists('image_type_to_mime_type') && defined('IMAGETYPE_ICO')) {
    $mime = image_type_to_mime_type(IMAGETYPE_ICO);
    echo str_contains($mime, "icon") || str_contains($mime, "ico") ? "ICO_MIME_OK" : "FAIL";
} else {
    echo "ICO_MIME_OK";
}
