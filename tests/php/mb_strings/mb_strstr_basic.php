<?php
// vybe-test: php/mb_strings/mb_strstr_basic
// origin: languages/php/tests/php/test_mb_strings.rs
// vybe-test-mode: compile

$s = "user@example.com";
echo mb_strstr($s, "@");       // @example.com
echo mb_strstr($s, "@", true); // user
