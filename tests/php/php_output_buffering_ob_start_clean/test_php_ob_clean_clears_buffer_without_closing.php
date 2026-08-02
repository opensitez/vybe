<?php
// vybe-test: php/php_output_buffering_ob_start_clean/test_php_ob_clean_clears_buffer_without_closing
// origin: languages/php/tests/php/test_php_output_buffering_ob_start_clean.rs
// vybe-test-mode: compile

ob_start();
echo "discarded text";
ob_clean();
echo "kept text";
$output = ob_get_clean();
echo $output;
