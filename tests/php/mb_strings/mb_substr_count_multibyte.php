<?php
// vybe-test: php/mb_strings/mb_substr_count_multibyte
// origin: languages/php/tests/php/test_mb_strings.rs
// vybe-test-mode: compile

$s = "日本日本日";
echo mb_substr_count($s, "日本");  // 2
echo mb_substr_count($s, "日");    // 3
