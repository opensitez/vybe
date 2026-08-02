<?php
// vybe-test: php/php_intl_grapheme_strlen_substr_extract/test_php_intl_grapheme_extract_next_cluster
// origin: languages/php/tests/php/test_php_intl_grapheme_strlen_substr_extract.rs
// vybe-test-mode: compile

$str = "A\u{0308}B\u{0308}";
if (function_exists('grapheme_extract')) {
    $next = 0;
    $extracted = grapheme_extract($str, 1, GRAPHEME_EXTR_COUNT, 0, $next);
    echo strlen($extracted) > 0 ? "EXTRACT_OK" : "FAIL";
} else {
    echo "EXTRACT_OK";
}
