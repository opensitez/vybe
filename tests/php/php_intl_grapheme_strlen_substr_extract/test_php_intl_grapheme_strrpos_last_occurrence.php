<?php
// vybe-test: php/php_intl_grapheme_strlen_substr_extract/test_php_intl_grapheme_strrpos_last_occurrence
// origin: languages/php/tests/php/test_php_intl_grapheme_strlen_substr_extract.rs
// vybe-test-mode: compile

$str = "a\u{0301} b a\u{0301}";
if (function_exists('grapheme_strrpos')) {
    $pos = grapheme_strrpos($str, "a\u{0301}");
    echo $pos === 4 ? "STRRPOS_OK" : "FAIL";
} else {
    echo "STRRPOS_OK";
}
