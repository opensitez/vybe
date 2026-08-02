<?php
// vybe-test: php/mb_strings/mb_substr_negative
// origin: languages/php/tests/php/test_mb_strings.rs
// vybe-test-mode: compile

$s = "Hello World";
echo mb_substr($s, -5);     // World
echo mb_substr($s, -5, 3);  // Wor
