<?php
// vybe-test: php/output_buffering/ob_implicit_flush_controls_runtime_behavior
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start();
echo 'x';
ob_implicit_flush(true);
ob_end_flush();
ob_start();
echo 'y';
ob_end_clean();
echo 'done';
