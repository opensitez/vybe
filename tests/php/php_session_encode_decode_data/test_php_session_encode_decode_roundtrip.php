<?php
// vybe-test: php/php_session_encode_decode_data/test_php_session_encode_decode_roundtrip
// origin: languages/php/tests/php/test_php_session_encode_decode_data.rs
// vybe-test-mode: compile

@session_start();
$_SESSION["items"] = ["item1", "item2"];
$encoded = @session_encode();

$_SESSION = [];
@session_decode($encoded);

echo count($_SESSION["items"] ?? []) === 2 ? "ROUNDTRIP_OK" : "FAIL";
@session_write_close();
