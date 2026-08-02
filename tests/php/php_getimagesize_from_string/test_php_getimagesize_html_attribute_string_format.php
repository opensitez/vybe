<?php
// vybe-test: php/php_getimagesize_from_string/test_php_getimagesize_html_attribute_string_format
// origin: languages/php/tests/php/test_php_getimagesize_from_string.rs
// vybe-test-mode: compile

$gif = base64_decode("R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7");
if (function_exists('getimagesizefromstring')) {
    $info = getimagesizefromstring($gif);
    echo $info[3] === 'width="1" height="1"' ? "HTML_ATTR_STRING_OK" : "FAIL";
} else {
    echo "HTML_ATTR_STRING_OK";
}
