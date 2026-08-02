<?php
// vybe-test: php/php_mbstring_multibyte_utf8_ops/test_php_mb_detect_encoding_auto
// origin: languages/php/tests/php/test_php_mbstring_multibyte_utf8_ops.rs
// vybe-test-mode: compile

$text = "Simple ASCII string";
$encoding = mb_detect_encoding($text, ["UTF-8", "ASCII", "ISO-8859-1"]);
echo $encoding;
