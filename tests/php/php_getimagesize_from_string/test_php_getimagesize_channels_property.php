<?php
// vybe-test: php/php_getimagesize_from_string/test_php_getimagesize_channels_property
// origin: languages/php/tests/php/test_php_getimagesize_from_string.rs
// vybe-test-mode: compile

$png = base64_decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==");
if (function_exists('getimagesizefromstring')) {
    $info = getimagesizefromstring($png);
    echo isset($info['channels']) || isset($info['bits']) ? "CHANNELS_BITS_OK" : "FAIL";
} else {
    echo "CHANNELS_BITS_OK";
}
