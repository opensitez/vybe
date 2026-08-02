<?php
// vybe-test: php/php_session_encode_decode_data/test_php_session_encode_boolean_and_null_types
// origin: languages/php/tests/php/test_php_session_encode_decode_data.rs
// vybe-test-mode: compile

@session_start();
$_SESSION["flag"] = true;
$_SESSION["nil"] = null;
$encoded = @session_encode();
$_SESSION = [];
@session_decode($encoded);
echo $_SESSION["flag"] === true && $_SESSION["nil"] === null ? "BOOL_NULL_DECODE_OK" : "FAIL";
@session_write_close();
