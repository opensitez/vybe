<?php
// vybe-test: php/function_builtins/register_shutdown_function_basic
// origin: languages/php/tests/php/test_function_builtins.rs
// vybe-test-mode: compile

register_shutdown_function(function() {
    echo 'shutdown';
});
echo 'running';
