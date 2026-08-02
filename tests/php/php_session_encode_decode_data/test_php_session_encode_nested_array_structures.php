<?php
// vybe-test: php/php_session_encode_decode_data/test_php_session_encode_nested_array_structures
// origin: languages/php/tests/php/test_php_session_encode_decode_data.rs
// vybe-test-mode: compile

@session_start();
$_SESSION["config"] = ["db" => ["host" => "localhost", "port" => 3306]];
$encoded = @session_encode();
$_SESSION = [];
@session_decode($encoded);
echo $_SESSION["config"]["db"]["port"] === 3306 ? "NESTED_ENCODE_OK" : "FAIL";
@session_write_close();
