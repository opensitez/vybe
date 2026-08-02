<?php
// vybe-test: php/mb_strings/mb_strpos_multibyte
// origin: languages/php/tests/php/test_mb_strings.rs
// vybe-test-mode: compile

$s = "こんにちは世界";
echo mb_strpos($s, "世界");  // 5
echo mb_strpos($s, "に");    // 2
