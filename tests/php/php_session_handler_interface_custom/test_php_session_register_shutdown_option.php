<?php
// vybe-test: php/php_session_handler_interface_custom/test_php_session_register_shutdown_option
// origin: languages/php/tests/php/test_php_session_handler_interface_custom.rs
// vybe-test-mode: compile

$sh = new SessionHandler();
$res = @session_set_save_handler($sh, true); // true = register_shutdown
echo $res !== null ? "REGISTER_SHUTDOWN_OK" : "FAIL";
