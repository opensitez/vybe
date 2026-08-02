<?php
// vybe-test: php/mb_strings/mb_strlen_multibyte
// origin: languages/php/tests/php/test_mb_strings.rs
// vybe-test-mode: compile

echo mb_strlen("héllo");      // 5 characters
echo mb_strlen("日本語");       // 3 characters
echo mb_strlen("emoji😀");     // 6 characters
