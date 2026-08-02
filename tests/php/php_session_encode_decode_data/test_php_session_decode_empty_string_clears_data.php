<?php
// vybe-test: php/php_session_encode_decode_data/test_php_session_decode_empty_string_clears_data
// origin: languages/php/tests/php/test_php_session_encode_decode_data.rs
// vybe-test-mode: compile

@session_start();
$_SESSION["data"] = 123;
@session_decode("");
echo count($_SESSION) === 0 ? "DECODE_EMPTY_CLEARS_OK" : "FAIL";
@session_write_close();
