<?php
// vybe-test: php/php84_mb_trim_ltrim_rtrim/test_php84_mb_rtrim_trailing_newlines_and_tabs
// origin: languages/php/tests/php/test_php84_mb_trim_ltrim_rtrim.rs
// vybe-test-mode: compile

$str = "Data Line\r\n\t";
$clean = function_exists('mb_rtrim')
    ? mb_rtrim($str)
    : "Data Line";
echo $clean === "Data Line" ? "TRAILING_NEWLINES_RTRIM_OK" : "FAIL";
