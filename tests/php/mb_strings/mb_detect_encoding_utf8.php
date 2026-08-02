<?php
// vybe-test: php/mb_strings/mb_detect_encoding_utf8
// origin: languages/php/tests/php/test_mb_strings.rs
// vybe-test-mode: compile

$s = "Hello World";
$enc = mb_detect_encoding($s, ['UTF-8', 'ASCII', 'ISO-8859-1']);
echo $enc !== false ? 'detected' : 'not detected';
