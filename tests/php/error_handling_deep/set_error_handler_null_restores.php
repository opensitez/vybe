<?php
// vybe-test: php/error_handling_deep/set_error_handler_null_restores
// origin: languages/php/tests/php/test_error_handling_deep.rs
// vybe-test-mode: compile

set_error_handler(fn() => true);
$prev = set_error_handler(null); // restores default
echo 'restored';
