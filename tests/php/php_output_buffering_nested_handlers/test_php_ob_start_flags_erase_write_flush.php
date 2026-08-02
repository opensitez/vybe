<?php
// vybe-test: php/php_output_buffering_nested_handlers/test_php_ob_start_flags_erase_write_flush
// origin: languages/php/tests/php/test_php_output_buffering_nested_handlers.rs
// vybe-test-mode: compile

// Buffer that cannot be cleaned, only flushed
ob_start(flags: PHP_OUTPUT_HANDLER_FLUSHALBLE | PHP_OUTPUT_HANDLER_REMOVABLE);
echo "non_erasable_content";
ob_end_flush();
