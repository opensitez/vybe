<?php
// vybe-test: php/php_intl_grapheme_strlen_substr_extract/test_php_intl_grapheme_stripos_case_insensitive
// origin: languages/php/tests/php/test_php_intl_grapheme_strlen_substr_extract.rs
// vybe-test-mode: compile

$str = "E\u{0301}xample";
if (function_exists('grapheme_stripos')) {
    $pos = grapheme_stripos($str, "e\u{0301}");
    echo $pos === 0 ? "STRIPOS_OK" : "FAIL";
} else {
    echo "STRIPOS_OK";
}
