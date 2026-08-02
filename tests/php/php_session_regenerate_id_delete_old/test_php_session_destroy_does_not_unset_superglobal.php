<?php
// vybe-test: php/php_session_regenerate_id_delete_old/test_php_session_destroy_does_not_unset_superglobal
// origin: languages/php/tests/php/test_php_session_regenerate_id_delete_old.rs
// vybe-test-mode: compile

@session_start();
$_SESSION["key"] = "value";
@session_destroy();
// Note: session_destroy() deletes session storage, but $_SESSION remains set in script memory until unset()
echo isset($_SESSION["key"]) ? "DESTROY_SUPERGLOBAL_PERSISTS_OK" : "FAIL";
