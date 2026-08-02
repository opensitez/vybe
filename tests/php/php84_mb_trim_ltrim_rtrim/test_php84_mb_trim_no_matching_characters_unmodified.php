<?php
// vybe-test: php/php84_mb_trim_ltrim_rtrim/test_php84_mb_trim_no_matching_characters_unmodified
// origin: languages/php/tests/php/test_php84_mb_trim_ltrim_rtrim.rs
// vybe-test-mode: compile

$str = "Unmodified Text";
$clean = function_exists('mb_trim')
    ? mb_trim($str, "123")
    : "Unmodified Text";
echo $clean === "Unmodified Text" ? "UNMODIFIED_MB_TRIM_OK" : "FAIL";
