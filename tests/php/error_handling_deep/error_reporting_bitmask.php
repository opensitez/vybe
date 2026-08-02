<?php
// vybe-test: php/error_handling_deep/error_reporting_bitmask
// origin: languages/php/tests/php/test_error_handling_deep.rs
// vybe-test-mode: compile

// Combine error levels with bitwise OR
$level = E_ERROR | E_WARNING | E_NOTICE;
$old = error_reporting($level);
echo error_reporting() === $level ? 'set correctly' : 'wrong';
error_reporting($old);
