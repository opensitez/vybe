<?php
// vybe-test: php/php_session_encode_decode_data/test_php_session_get_cookie_params_structure
// origin: languages/php/tests/php/test_php_session_encode_decode_data.rs
// vybe-test-mode: compile

$params = session_get_cookie_params();
echo isset($params["lifetime"]) && isset($params["path"]) && isset($params["domain"]) ? "COOKIE_PARAMS_OK" : "FAIL";
