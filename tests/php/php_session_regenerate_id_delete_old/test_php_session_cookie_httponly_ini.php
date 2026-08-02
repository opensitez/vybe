<?php
// vybe-test: php/php_session_regenerate_id_delete_old/test_php_session_cookie_httponly_ini
// origin: languages/php/tests/php/test_php_session_regenerate_id_delete_old.rs
// vybe-test-mode: compile

$httponly = ini_get("session.cookie_httponly");
echo $httponly !== false ? "HTTPONLY_INI_OK" : "FAIL";
