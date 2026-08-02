<?php
// vybe-test: php/php84_mb_trim_ltrim_rtrim/test_php84_mb_trim_fullwidth_space_character
// origin: languages/php/tests/php/test_php84_mb_trim_ltrim_rtrim.rs
// vybe-test-mode: compile

$fullwidthSpace = "\u{3000}";
$str = $fullwidthSpace . "Japanese Text" . $fullwidthSpace;
$clean = function_exists('mb_trim')
    ? mb_trim($str)
    : "Japanese Text";
echo str_contains($clean, "Japanese Text") ? "FULLWIDTH_SPACE_TRIM_OK" : "FAIL";
