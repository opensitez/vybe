<?php
// vybe-test: php/php_session_regenerate_id_delete_old/test_php_session_gc_maxlifetime_ini
// origin: languages/php/tests/php/test_php_session_regenerate_id_delete_old.rs
// vybe-test-mode: compile

$ttl = ini_get("session.gc_maxlifetime");
echo is_numeric($ttl) && (int)$ttl > 0 ? "GC_MAXLIFETIME_INI_OK" : "FAIL";
