<?php
// vybe-test: php/php_proc_terminate_signal_handling/test_php_signal_constants_defined
// origin: languages/php/tests/php/test_php_proc_terminate_signal_handling.rs
// vybe-test-mode: compile

$hasSig = defined('SIGTERM') && defined('SIGKILL') && defined('SIGINT');
echo $hasSig ? "SIGNAL_CONSTANTS_DEFINED" : "FAIL";
