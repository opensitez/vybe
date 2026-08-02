<?php
// vybe-test: php/output_buffering/ob_end_flush
// origin: languages/php/tests/php/test_output_buffering.rs
// vybe-test-mode: compile

ob_start();
echo "flushed";
ob_end_flush();
