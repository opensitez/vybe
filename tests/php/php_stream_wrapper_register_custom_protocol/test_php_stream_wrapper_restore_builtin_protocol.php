<?php
// vybe-test: php/php_stream_wrapper_register_custom_protocol/test_php_stream_wrapper_restore_builtin_protocol
// origin: languages/php/tests/php/test_php_stream_wrapper_register_custom_protocol.rs
// vybe-test-mode: compile

stream_wrapper_unregister("file");
echo !in_array("file", stream_get_wrappers()) ? "UNREGISTERED_FILE" : "FAIL";
stream_wrapper_restore("file");
echo in_array("file", stream_get_wrappers()) ? " RESTORED_FILE" : " FAIL";
