<?php
// vybe-test: php/php_ob_list_handlers_status/test_php_ob_get_status_buffer_used_property
// origin: languages/php/tests/php/test_php_ob_list_handlers_status.rs
// vybe-test-mode: compile

ob_start();
echo "12345";
$status = ob_get_status(false);
ob_end_clean();
echo $status["buffer_used"] === 5 ? "BUFFER_USED_5_OK" : "FAIL";
