<?php
// vybe-test: php/php_ob_list_handlers_status/test_php_ob_list_handlers_empty_stack
// origin: languages/php/tests/php/test_php_ob_list_handlers_status.rs
// vybe-test-mode: compile

while (ob_get_level() > 0) ob_end_clean();
$handlers = ob_list_handlers();
echo is_array($handlers) && count($handlers) === 0 ? "EMPTY_HANDLERS_STACK_OK" : "FAIL";
