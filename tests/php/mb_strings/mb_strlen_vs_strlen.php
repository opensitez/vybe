<?php
// vybe-test: php/mb_strings/mb_strlen_vs_strlen
// origin: languages/php/tests/php/test_mb_strings.rs
// vybe-test-mode: compile

$s = "café";
$byte_len = strlen($s);      // bytes
$char_len = mb_strlen($s);   // characters
echo $char_len . ':' . ($byte_len >= $char_len ? 'bytes>=chars' : 'fail');
