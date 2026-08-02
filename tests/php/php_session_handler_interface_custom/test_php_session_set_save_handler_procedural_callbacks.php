<?php
// vybe-test: php/php_session_handler_interface_custom/test_php_session_set_save_handler_procedural_callbacks
// origin: languages/php/tests/php/test_php_session_handler_interface_custom.rs
// vybe-test-mode: compile

$res = session_set_save_handler(
    fn($path, $name) => true, // open
    fn() => true,             // close
    fn($id) => "",            // read
    fn($id, $data) => true,   // write
    fn($id) => true,          // destroy
    fn($max) => 0             // gc
);
echo $res ? "PROCEDURAL_SAVE_HANDLER_OK" : "FAIL";
