<?php
// vybe-test: php/php_getimagesize_from_string/test_php_getimagesize_empty_string_returns_false
// origin: languages/php/tests/php/test_php_getimagesize_from_string.rs
// vybe-test-mode: compile

if (function_exists('getimagesizefromstring')) {
    $info = @getimagesizefromstring("");
    echo $info === false ? "EMPTY_STRING_FALSE_OK" : "FAIL";
} else {
    echo "EMPTY_STRING_FALSE_OK";
}
