<?php
// vybe-test: php/php_output_buffering_nested_handlers/test_php_ob_get_length_buffer_byte_count
// origin: languages/php/tests/php/test_php_output_buffering_nested_handlers.rs

ob_start();
echo "1234567890";
$len = ob_get_length();
ob_end_clean();
echo "Length: $len";
