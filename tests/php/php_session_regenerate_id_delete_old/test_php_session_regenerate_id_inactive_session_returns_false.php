<?php
// vybe-test: php/php_session_regenerate_id_delete_old/test_php_session_regenerate_id_inactive_session_returns_false
// origin: languages/php/tests/php/test_php_session_regenerate_id_delete_old.rs
// vybe-test-mode: compile

$res = @session_regenerate_id(false);
echo $res === false ? "INACTIVE_REGENERATE_FALSE" : "FAIL";
