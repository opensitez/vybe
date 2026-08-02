<?php
// vybe-test: php/php_session_encode_decode_data/test_php_session_abort_discards_changes
// origin: languages/php/tests/php/test_php_session_encode_decode_data.rs
// vybe-test-mode: compile

@session_start();
$_SESSION["temp"] = "discard_me";
@session_abort();
echo session_status() === PHP_SESSION_NONE ? "ABORT_STATUS_NONE_OK" : "FAIL";
