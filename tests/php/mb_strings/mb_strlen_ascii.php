<?php
// vybe-test: php/mb_strings/mb_strlen_ascii
// origin: languages/php/tests/php/test_mb_strings.rs
// vybe-test-mode: compile

echo mb_strlen("hello");
echo mb_strlen("hello", 'UTF-8');
