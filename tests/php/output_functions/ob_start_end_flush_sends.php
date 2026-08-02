<?php
// vybe-test: php/output_functions/ob_start_end_flush_sends
// origin: languages/php/tests/php/test_output_functions.rs
// vybe-test-mode: compile

ob_start();
echo 'flushed content';
ob_end_flush();
