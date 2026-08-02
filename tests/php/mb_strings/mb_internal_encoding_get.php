<?php
// vybe-test: php/mb_strings/mb_internal_encoding_get
// origin: languages/php/tests/php/test_mb_strings.rs
// vybe-test-mode: compile

$enc = mb_internal_encoding();
echo is_string($enc) ? 'is string' : 'not string';
