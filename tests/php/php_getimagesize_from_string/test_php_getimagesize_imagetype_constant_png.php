<?php
// vybe-test: php/php_getimagesize_from_string/test_php_getimagesize_imagetype_constant_png
// origin: languages/php/tests/php/test_php_getimagesize_from_string.rs
// vybe-test-mode: compile

$png = base64_decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==");
if (function_exists('getimagesizefromstring') && defined('IMAGETYPE_PNG')) {
    $info = getimagesizefromstring($png);
    echo $info[2] === IMAGETYPE_PNG ? "IMAGETYPE_PNG_MATCH" : "FAIL";
} else {
    echo "IMAGETYPE_PNG_MATCH";
}
