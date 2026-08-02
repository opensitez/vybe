<?php
// vybe-test: php/output_buffering/ob_start_with_callback
// origin: languages/php/tests/php/test_output_buffering.rs
// vybe-test-mode: compile

ob_start(fn(string $buf) => strtoupper($buf));
echo "hello world";
ob_end_flush();
