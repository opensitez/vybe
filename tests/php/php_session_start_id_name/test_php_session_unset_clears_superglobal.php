<?php
// vybe-test: php/php_session_start_id_name/test_php_session_unset_clears_superglobal
// origin: languages/php/tests/php/test_php_session_start_id_name.rs
// vybe-test-mode: compile

@session_start();
$_SESSION["user"] = "Alice";
@session_unset();
echo count($_SESSION) === 0 ? "UNSET_OK" : "FAIL";
@session_write_close();
