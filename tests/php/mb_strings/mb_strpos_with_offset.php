<?php
// vybe-test: php/mb_strings/mb_strpos_with_offset
// origin: languages/php/tests/php/test_mb_strings.rs
// vybe-test-mode: compile

$s = "abcabc";
echo mb_strpos($s, "b", 0);  // 1
echo mb_strpos($s, "b", 2);  // 4
