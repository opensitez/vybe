<?php
// vybe-test: php/php_mbstring_multibyte_utf8_ops/test_php_mb_convert_encoding_conversion
// origin: languages/php/tests/php/test_php_mbstring_multibyte_utf8_ops.rs
// vybe-test-mode: compile

$utf8 = "Test string";
$iso = mb_convert_encoding($utf8, "ISO-8859-1", "UTF-8");
echo strlen($iso) > 0 ? "CONVERTED" : "FAIL";
