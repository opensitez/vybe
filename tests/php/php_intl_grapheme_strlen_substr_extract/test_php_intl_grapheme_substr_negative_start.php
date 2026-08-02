<?php
// vybe-test: php/php_intl_grapheme_strlen_substr_extract/test_php_intl_grapheme_substr_negative_start
// origin: languages/php/tests/php/test_php_intl_grapheme_strlen_substr_extract.rs
// vybe-test-mode: compile

$str = "A\u{0301}B\u{0301}C\u{0301}";
if (function_exists('grapheme_substr')) {
    $last = grapheme_substr($str, -1);
    echo grapheme_strlen($last) === 1 ? "NEGATIVE_SUBSTR_OK" : "FAIL";
} else {
    echo "NEGATIVE_SUBSTR_OK";
}
