<?php
// vybe-test: php/mb_strings/mb_strrpos_basic
// origin: languages/php/tests/php/test_mb_strings.rs
// vybe-test-mode: compile

$s = "hello world hello";
echo mb_strrpos($s, "hello");  // 12
echo mb_strrpos($s, "o");
