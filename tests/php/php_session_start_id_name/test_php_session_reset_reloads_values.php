<?php
// vybe-test: php/php_session_start_id_name/test_php_session_reset_reloads_values
// origin: languages/php/tests/php/test_php_session_start_id_name.rs
// vybe-test-mode: compile

@session_start();
$_SESSION["key"] = "original";
@session_reset();
echo session_status() === PHP_SESSION_ACTIVE ? "RESET_OK" : "FAIL";
@session_write_close();
