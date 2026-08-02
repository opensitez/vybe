<?php
// vybe-test: php/mb_strings/mb_detect_encoding_multibyte
// origin: languages/php/tests/php/test_mb_strings.rs
// vybe-test-mode: compile

$s = "こんにちは";
$enc = mb_detect_encoding($s, 'auto');
echo $enc !== false ? 'detected' : 'not detected';
