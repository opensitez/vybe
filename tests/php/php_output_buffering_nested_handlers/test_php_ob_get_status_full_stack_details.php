<?php
// vybe-test: php/php_output_buffering_nested_handlers/test_php_ob_get_status_full_stack_details
// origin: languages/php/tests/php/test_php_output_buffering_nested_handlers.rs
// vybe-test-mode: compile

ob_start();
$statuses = ob_get_status(full_status: true);
echo is_array($statuses) && count($statuses) > 0 ? "STATUS_OK" : "FAIL";
ob_end_clean();
