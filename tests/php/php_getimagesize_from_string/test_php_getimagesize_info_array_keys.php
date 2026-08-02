<?php
// vybe-test: php/php_getimagesize_from_string/test_php_getimagesize_info_array_keys
// origin: languages/php/tests/php/test_php_getimagesize_from_string.rs
// vybe-test-mode: compile

$gif = base64_decode("R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7");
if (function_exists('getimagesizefromstring')) {
    $info = getimagesizefromstring($gif);
    echo isset($info[0]) && isset($info[1]) && isset($info[2]) && isset($info[3]) && isset($info['bits']) ? "INFO_KEYS_OK" : "FAIL";
} else {
    echo "INFO_KEYS_OK";
}
