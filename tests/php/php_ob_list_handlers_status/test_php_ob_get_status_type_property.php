<?php
// vybe-test: php/php_ob_list_handlers_status/test_php_ob_get_status_type_property
// origin: languages/php/tests/php/test_php_ob_list_handlers_status.rs
// vybe-test-mode: compile

ob_start();
$status = ob_get_status(false);
ob_end_clean();
echo isset($status["type"]) && is_int($status["type"]) ? "STATUS_TYPE_INT_OK" : "FAIL";
