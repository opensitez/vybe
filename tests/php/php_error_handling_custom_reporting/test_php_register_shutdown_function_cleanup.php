<?php
// vybe-test: php/php_error_handling_custom_reporting/test_php_register_shutdown_function_cleanup
// origin: languages/php/tests/php/test_php_error_handling_custom_reporting.rs
// vybe-test-mode: compile

register_shutdown_function(function() {
    // Shutdown cleanup callback
});
