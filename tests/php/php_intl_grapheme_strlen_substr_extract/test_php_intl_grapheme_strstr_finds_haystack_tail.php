<?php
// vybe-test: php/php_intl_grapheme_strlen_substr_extract/test_php_intl_grapheme_strstr_finds_haystack_tail
// origin: languages/php/tests/php/test_php_intl_grapheme_strlen_substr_extract.rs
// vybe-test-mode: compile

$str = "alpha\u{0301}beta";
if (function_exists('grapheme_strstr')) {
    $tail = grapheme_strstr($str, "b");
    echo $tail === "beta" ? "STRSTR_OK" : "FAIL";
} else {
    echo "STRSTR_OK";
}
