<?php
// vybe-test: php/output_buffering/ob_end_flush_with_no_handler
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start(function(string $buf): string { return $buf . 'X'; });
echo 'A';
echo ob_end_flush();
