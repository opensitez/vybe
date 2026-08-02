<?php
// vybe-test: php/php_session_start_id_name/test_php_session_start_options_read_and_close
// origin: languages/php/tests/php/test_php_session_start_id_name.rs
// vybe-test-mode: compile

$status = @session_start([
    "read_and_close" => true,
    "cookie_lifetime" => 3600
]);
echo $status && session_status() === PHP_SESSION_NONE ? "READ_AND_CLOSE_OK" : "FAIL";
