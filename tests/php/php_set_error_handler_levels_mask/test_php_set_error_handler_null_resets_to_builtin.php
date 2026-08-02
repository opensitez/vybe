<?php
// vybe-test: php/php_set_error_handler_levels_mask/test_php_set_error_handler_null_resets_to_builtin
// origin: languages/php/tests/php/test_php_set_error_handler_levels_mask.rs
// vybe-test-mode: compile

set_error_handler(fn() => true);
set_error_handler(null);
echo "NULL_RESET_OK";
