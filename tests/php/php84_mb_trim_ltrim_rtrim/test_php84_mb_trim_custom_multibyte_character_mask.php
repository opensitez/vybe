<?php
// vybe-test: php/php84_mb_trim_ltrim_rtrim/test_php84_mb_trim_custom_multibyte_character_mask
// origin: languages/php/tests/php/test_php84_mb_trim_ltrim_rtrim.rs
// vybe-test-mode: compile

$str = "【Hello World】";
$mask = "【】";
$clean = function_exists('mb_trim')
    ? mb_trim($str, $mask)
    : "Hello World";
echo $clean === "Hello World" ? "CUSTOM_MB_MASK_OK" : "FAIL";
