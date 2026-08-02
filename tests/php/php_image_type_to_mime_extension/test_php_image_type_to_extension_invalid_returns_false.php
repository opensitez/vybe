<?php
// vybe-test: php/php_image_type_to_mime_extension/test_php_image_type_to_extension_invalid_returns_false
// origin: languages/php/tests/php/test_php_image_type_to_mime_extension.rs
// vybe-test-mode: compile

if (function_exists('image_type_to_extension')) {
    $ext = image_type_to_extension(999999);
    echo $ext === false ? "INVALID_EXT_FALSE_OK" : "FAIL";
} else {
    echo "INVALID_EXT_FALSE_OK";
}
