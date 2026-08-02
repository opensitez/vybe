<?php
// vybe-test: php/php84_mb_trim_ltrim_rtrim/test_php84_mb_trim_empty_string
// origin: languages/php/tests/php/test_php84_mb_trim_ltrim_rtrim.rs
// vybe-test-mode: compile

$clean = function_exists('mb_trim')
    ? mb_trim("")
    : "";
echo $clean === "" ? "EMPTY_MB_TRIM_OK" : "FAIL";
