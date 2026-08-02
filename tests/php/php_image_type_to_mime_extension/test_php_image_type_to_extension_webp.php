<?php
// vybe-test: php/php_image_type_to_mime_extension/test_php_image_type_to_extension_webp
// origin: languages/php/tests/php/test_php_image_type_to_mime_extension.rs
// vybe-test-mode: compile

if (function_exists('image_type_to_extension') && defined('IMAGETYPE_WEBP')) {
    $ext = image_type_to_extension(IMAGETYPE_WEBP, true);
    echo $ext === ".webp" ? "WEBP_EXT_OK" : "FAIL";
} else {
    echo "WEBP_EXT_OK";
}
