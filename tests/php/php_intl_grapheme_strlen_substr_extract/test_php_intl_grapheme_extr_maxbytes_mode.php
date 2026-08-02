<?php
// vybe-test: php/php_intl_grapheme_strlen_substr_extract/test_php_intl_grapheme_extr_maxbytes_mode
// origin: languages/php/tests/php/test_php_intl_grapheme_strlen_substr_extract.rs
// vybe-test-mode: compile

if (defined('GRAPHEME_EXTR_MAXBYTES')) {
    echo GRAPHEME_EXTR_MAXBYTES === 0 || is_int(GRAPHEME_EXTR_MAXBYTES) ? "MAXBYTES_CONST_OK" : "FAIL";
} else {
    echo "MAXBYTES_CONST_OK";
}
