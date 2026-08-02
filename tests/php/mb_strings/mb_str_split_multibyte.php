<?php
// vybe-test: php/mb_strings/mb_str_split_multibyte
// origin: languages/php/tests/php/test_mb_strings.rs
// vybe-test-mode: compile

$chars = mb_str_split("日本語");
echo count($chars) . ':' . $chars[0] . $chars[1] . $chars[2];
