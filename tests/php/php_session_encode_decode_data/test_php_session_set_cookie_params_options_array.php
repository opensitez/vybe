<?php
// vybe-test: php/php_session_encode_decode_data/test_php_session_set_cookie_params_options_array
// origin: languages/php/tests/php/test_php_session_encode_decode_data.rs
// vybe-test-mode: compile

session_set_cookie_params([
    "lifetime" => 7200,
    "path" => "/app",
    "secure" => true,
    "httponly" => true,
    "samesite" => "Lax"
]);
$p = session_get_cookie_params();
echo $p["lifetime"] === 7200 && $p["path"] === "/app" ? "SET_COOKIE_PARAMS_OK" : "FAIL";
