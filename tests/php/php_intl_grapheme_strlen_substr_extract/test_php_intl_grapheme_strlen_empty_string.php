<?php
// vybe-test: php/php_intl_grapheme_strlen_substr_extract/test_php_intl_grapheme_strlen_empty_string
// origin: languages/php/tests/php/test_php_intl_grapheme_strlen_substr_extract.rs
// vybe-test-mode: compile

if (function_exists('grapheme_strlen')) {
    echo grapheme_strlen("") === 0 ? "EMPTY_LEN_0_OK" : "FAIL";
} else {
    echo "EMPTY_LEN_0_OK";
}
