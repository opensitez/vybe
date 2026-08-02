<?php
// vybe-test: php/php_getimagesize_from_string/test_php_getimagesizefromstring_info_array_iptc_capture
// origin: languages/php/tests/php/test_php_getimagesize_from_string.rs
// vybe-test-mode: compile

$gif = base64_decode("R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7");
if (function_exists('getimagesizefromstring')) {
    $info = getimagesizefromstring($gif, $imageInfo);
    echo is_array($info) ? "IPTC_INFO_PARAM_OK" : "FAIL";
} else {
    echo "IPTC_INFO_PARAM_OK";
}
