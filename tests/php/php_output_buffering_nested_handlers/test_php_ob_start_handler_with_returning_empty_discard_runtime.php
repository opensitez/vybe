<?php
// vybe-test: php/php_output_buffering_nested_handlers/test_php_ob_start_handler_with_returning_empty_discard_runtime
// origin: languages/php/tests/php/test_php_output_buffering_nested_handlers.rs

ob_start(function(string $chunk): string {
    return '';
});
echo 'should disappear';
ob_end_flush();
echo 'after';
