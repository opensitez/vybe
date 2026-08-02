<?php
// vybe-test: php/php_output_buffering_nested_handlers/test_php_ob_start_multiple_custom_filters
// origin: languages/php/tests/php/test_php_output_buffering_nested_handlers.rs
// vybe-test-mode: compile

ob_start(fn($s) => str_replace("foo", "bar", $s));
ob_start(fn($s) => strtoupper($s));
echo "foo text";
ob_end_flush();
ob_end_flush();
