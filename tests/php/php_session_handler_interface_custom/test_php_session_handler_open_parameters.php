<?php
// vybe-test: php/php_session_handler_interface_custom/test_php_session_handler_open_parameters
// origin: languages/php/tests/php/test_php_session_handler_interface_custom.rs
// vybe-test-mode: compile

$openedPath = "";
$openedName = "";
session_set_save_handler(
    function($path, $name) use (&$openedPath, &$openedName) {
        $openedPath = $path;
        $openedName = $name;
        return true;
    },
    fn() => true, fn($i) => "", fn($i, $d) => true, fn($i) => true, fn($m) => 0
);
@session_start();
@session_write_close();
echo "OPEN_PARAMS_OK";
