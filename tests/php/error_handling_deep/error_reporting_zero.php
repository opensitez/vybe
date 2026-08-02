<?php
// vybe-test: php/error_handling_deep/error_reporting_zero
// origin: languages/php/tests/php/test_error_handling_deep.rs
// vybe-test-mode: compile

$old = error_reporting(0);
echo error_reporting() === 0 ? 'suppressed' : 'not suppressed';
error_reporting($old);
