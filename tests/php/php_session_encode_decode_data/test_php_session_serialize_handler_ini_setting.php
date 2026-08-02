<?php
// vybe-test: php/php_session_encode_decode_data/test_php_session_serialize_handler_ini_setting
// origin: languages/php/tests/php/test_php_session_encode_decode_data.rs
// vybe-test-mode: compile

$handler = ini_get("session.serialize_handler");
echo is_string($handler) && strlen($handler) > 0 ? "SERIALIZE_HANDLER_INI_OK" : "FAIL";
