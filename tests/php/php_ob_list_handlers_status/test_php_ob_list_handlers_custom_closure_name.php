<?php
// vybe-test: php/php_ob_list_handlers_status/test_php_ob_list_handlers_custom_closure_name
// origin: languages/php/tests/php/test_php_ob_list_handlers_status.rs
// vybe-test-mode: compile

ob_start(fn($s) => $s);
$handlers = ob_list_handlers();
ob_end_clean();
echo str_contains($handlers[0], "Closure") || str_contains($handlers[0], "default") ? "CLOSURE_HANDLER_NAME_OK" : "FAIL";
