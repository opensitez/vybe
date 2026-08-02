<?php
// vybe-test: php/php_session_handler_interface_custom/test_php_session_gc_garbage_collection_invocation
// origin: languages/php/tests/php/test_php_session_handler_interface_custom.rs
// vybe-test-mode: compile

if (function_exists('session_gc')) {
    $collected = @session_gc();
    echo is_int($collected) || $collected === false ? "SESSION_GC_OK" : "FAIL";
} else {
    echo "SESSION_GC_OK";
}
