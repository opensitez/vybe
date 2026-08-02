<?php
// vybe-test: php/mb_strings/mb_convert_encoding_latin_to_utf8
// origin: languages/php/tests/php/test_mb_strings.rs
// vybe-test-mode: compile

// Converting between compatible encodings
$s = "Hello";
$out = mb_convert_encoding($s, 'UTF-8', 'ASCII');
echo strlen($out) > 0 ? 'converted' : 'empty';
