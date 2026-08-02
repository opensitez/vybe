<?php
// vybe-test: php/mb_strings/mb_strlen_with_encoding_param
// origin: languages/php/tests/php/test_mb_strings.rs
// vybe-test-mode: compile

$s = "Hello";
echo mb_strlen($s, 'UTF-8');
echo mb_strlen($s, 'ASCII');
