<?php
// vybe-test: php/php_session_start_id_name/test_php_session_save_path_getter_setter
// origin: languages/php/tests/php/test_php_session_start_id_name.rs
// vybe-test-mode: compile

$orig = @session_save_path();
@session_save_path("/tmp");
$newPath = @session_save_path();
@session_save_path($orig);
echo $newPath === "/tmp" ? "SAVE_PATH_OK" : "FAIL";
