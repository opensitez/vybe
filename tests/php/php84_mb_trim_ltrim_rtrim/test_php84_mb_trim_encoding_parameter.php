<?php
// vybe-test: php/php84_mb_trim_ltrim_rtrim/test_php84_mb_trim_encoding_parameter
// origin: languages/php/tests/php/test_php84_mb_trim_ltrim_rtrim.rs
// vybe-test-mode: compile

$str = "   Multibyte Encoding   ";
$clean = function_exists('mb_trim')
    ? mb_trim($str, null, "UTF-8")
    : "Multibyte Encoding";
echo $clean === "Multibyte Encoding" ? "MB_ENCODING_PARAM_OK" : "FAIL";
