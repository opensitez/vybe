<?php
// vybe-test: php/mb_strings/mb_substr_multibyte
// origin: languages/php/tests/php/test_mb_strings.rs
// vybe-test-mode: compile

$s = "日本語テスト";
echo mb_substr($s, 0, 3);  // 日本語
echo mb_substr($s, 3);     // テスト
