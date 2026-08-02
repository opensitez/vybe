<?php
// vybe-test: php/php_getimagesize_from_string/test_php_getimagesize_imagetype_constant_gif
// origin: languages/php/tests/php/test_php_getimagesize_from_string.rs
// vybe-test-mode: compile

$gif = base64_decode("R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7");
if (function_exists('getimagesizefromstring') && defined('IMAGETYPE_GIF')) {
    $info = getimagesizefromstring($gif);
    echo $info[2] === IMAGETYPE_GIF ? "IMAGETYPE_GIF_MATCH" : "FAIL";
} else {
    echo "IMAGETYPE_GIF_MATCH";
}
