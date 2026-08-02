<?php
// vybe-test: php/output_buffering/ob_end_flush_returns_true_on_active_buffer
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start();
echo 'ok';
echo ob_end_flush();
