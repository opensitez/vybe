<?php
// vybe-test: php/mb_strings/mb_strpos_basic
// origin: languages/php/tests/php/test_mb_strings.rs
// vybe-test-mode: compile

$s = "Hello World";
echo mb_strpos($s, "World");
echo mb_strpos($s, "o");
var_dump(mb_strpos($s, "xyz"));
