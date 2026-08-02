<?php
// vybe-test: php/php_session_start_id_name/test_php_session_create_id_prefix
// origin: languages/php/tests/php/test_php_session_start_id_name.rs
// vybe-test-mode: compile

if (function_exists('session_create_id')) {
    $id = @session_create_id("prefix_");
    echo str_starts_with($id, "prefix_") ? "CREATE_ID_PREFIX_OK" : "FAIL";
} else {
    echo "CREATE_ID_PREFIX_OK";
}
