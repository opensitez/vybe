<?php
// vybe-test: php/php_mbstring_multibyte_utf8_ops/test_php_mb_str_split_chunking
// origin: languages/php/tests/php/test_php_mbstring_multibyte_utf8_ops.rs
// vybe-test-mode: compile

$str = "你好世界";
$chars = mb_str_split($str, 1, "UTF-8");
echo implode("-", $chars);
