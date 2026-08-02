<?php
// vybe-test: php/output_buffering/ob_end_flush_then_get_level_zero
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start();
echo 'x';
ob_end_flush();
echo ob_get_level();
