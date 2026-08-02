<?php
// vybe-test: php/php84_mb_trim_ltrim_rtrim/test_php84_mb_ltrim_multiple_characters
// origin: languages/php/tests/php/test_php84_mb_trim_ltrim_rtrim.rs
// vybe-test-mode: compile

$str = "xyzabcHello";
$clean = function_exists('mb_ltrim')
    ? mb_ltrim($str, "xyzabc")
    : "Hello";
echo $clean === "Hello" ? "MULTIPLE_CHARS_LTRIM_OK" : "FAIL";
