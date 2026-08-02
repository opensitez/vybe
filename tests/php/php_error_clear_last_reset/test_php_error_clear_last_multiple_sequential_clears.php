<?php
// vybe-test: php/php_error_clear_last_reset/test_php_error_clear_last_multiple_sequential_clears
// origin: languages/php/tests/php/test_php_error_clear_last_reset.rs
// vybe-test-mode: compile

error_clear_last();
error_clear_last();
error_clear_last();
echo error_get_last() === null ? "SEQUENTIAL_CLEARS_OK" : "FAIL";
