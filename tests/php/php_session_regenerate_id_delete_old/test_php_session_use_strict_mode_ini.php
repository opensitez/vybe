<?php
// vybe-test: php/php_session_regenerate_id_delete_old/test_php_session_use_strict_mode_ini
// origin: languages/php/tests/php/test_php_session_regenerate_id_delete_old.rs
// vybe-test-mode: compile

$strict = ini_get("session.use_strict_mode");
echo $strict !== false ? "USE_STRICT_MODE_INI_OK" : "FAIL";
