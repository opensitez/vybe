<?php
// vybe-test: php/magic_constants/magic_line_basic
// origin: languages/php/tests/php/test_magic_constants.rs
// vybe-test-mode: compile

$line = __LINE__;
echo is_int($line) ? 'is int' : 'not int';
echo $line > 0 ? ':positive' : ':zero';
