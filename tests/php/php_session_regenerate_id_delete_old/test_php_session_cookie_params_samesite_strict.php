<?php
// vybe-test: php/php_session_regenerate_id_delete_old/test_php_session_cookie_params_samesite_strict
// origin: languages/php/tests/php/test_php_session_regenerate_id_delete_old.rs
// vybe-test-mode: compile

session_set_cookie_params([
    "samesite" => "Strict",
    "secure" => true
]);
$p = session_get_cookie_params();
echo $p["samesite"] === "Strict" && $p["secure"] ? "SAMESITE_STRICT_OK" : "FAIL";
