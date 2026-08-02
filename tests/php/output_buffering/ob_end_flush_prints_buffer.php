<?php
// vybe-test: php/output_buffering/ob_end_flush_prints_buffer
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start();
echo 'flushme';
ob_end_flush();
