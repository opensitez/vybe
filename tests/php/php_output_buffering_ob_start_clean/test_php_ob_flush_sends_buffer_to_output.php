<?php
// vybe-test: php/php_output_buffering_ob_start_clean/test_php_ob_flush_sends_buffer_to_output
// origin: languages/php/tests/php/test_php_output_buffering_ob_start_clean.rs
// vybe-test-mode: compile

ob_start();
echo "flushed part 1\n";
ob_flush();
echo "flushed part 2\n";
ob_end_clean();
