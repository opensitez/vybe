<?php
// vybe-test: php/php_ob_list_handlers_status/test_php_ob_get_status_empty_stack_returns_empty_array
// origin: languages/php/tests/php/test_php_ob_list_handlers_status.rs
// vybe-test-mode: compile

while (ob_get_level() > 0) ob_end_clean();
$status = ob_get_status(true);
echo is_array($status) && count($status) === 0 ? "EMPTY_STATUS_STACK_OK" : "FAIL";
