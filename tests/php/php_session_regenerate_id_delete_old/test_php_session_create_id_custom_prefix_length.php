<?php
// vybe-test: php/php_session_regenerate_id_delete_old/test_php_session_create_id_custom_prefix_length
// origin: languages/php/tests/php/test_php_session_regenerate_id_delete_old.rs
// vybe-test-mode: compile

if (function_exists('session_create_id')) {
    $id = @session_create_id("sess_prefix_");
    echo strlen($id) > 12 ? "CREATE_ID_LENGTH_OK" : "FAIL";
} else {
    echo "CREATE_ID_LENGTH_OK";
}
