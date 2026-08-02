<?php
// vybe-test: php/php_stream_wrapper_register_custom_protocol/test_php_stream_wrapper_duplicate_register_error
// origin: languages/php/tests/php/test_php_stream_wrapper_register_custom_protocol.rs
// vybe-test-mode: compile

class DupWrapper {}
stream_wrapper_register("dupproto", DupWrapper::class);
$res = @stream_wrapper_register("dupproto", DupWrapper::class);
stream_wrapper_unregister("dupproto");
echo $res === false ? "DUPLICATE_REGISTER_FALSE" : "FAIL";
