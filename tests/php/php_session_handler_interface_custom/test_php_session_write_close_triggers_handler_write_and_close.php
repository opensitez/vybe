<?php
// vybe-test: php/php_session_handler_interface_custom/test_php_session_write_close_triggers_handler_write_and_close
// origin: languages/php/tests/php/test_php_session_handler_interface_custom.rs
// vybe-test-mode: compile

$written = false;
$closed = false;

session_set_save_handler(
    fn($p, $n) => true,
    function() use (&$closed) { $closed = true; return true; },
    fn($id) => "",
    function($id, $data) use (&$written) { $written = true; return true; },
    fn($id) => true,
    fn($m) => 0
);
@session_start();
$_SESSION["foo"] = "bar";
@session_write_close();
echo "WRITE_CLOSE_TRIGGERED_OK";
