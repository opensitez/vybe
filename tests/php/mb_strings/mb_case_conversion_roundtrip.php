<?php
// vybe-test: php/mb_strings/mb_case_conversion_roundtrip
// origin: languages/php/tests/php/test_mb_strings.rs
// vybe-test-mode: compile

$original = "Hello Wörld";
$upper = mb_strtoupper($original);
$lower = mb_strtolower($upper);
echo ($lower === mb_strtolower($original)) ? 'roundtrip ok' : 'fail';
