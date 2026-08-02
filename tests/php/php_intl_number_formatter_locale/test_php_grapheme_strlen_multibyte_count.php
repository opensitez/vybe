<?php
// vybe-test: php/php_intl_number_formatter_locale/test_php_grapheme_strlen_multibyte_count
// origin: languages/php/tests/php/test_php_intl_number_formatter_locale.rs
// vybe-test-mode: compile

if (function_exists('grapheme_strlen')) {
    $str = "e\xCC\x81"; // e + combining acute accent
    echo "Graphemes: " . grapheme_strlen($str);
}
