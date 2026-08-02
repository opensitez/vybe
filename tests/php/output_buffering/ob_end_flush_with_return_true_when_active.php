<?php
// vybe-test: php/output_buffering/ob_end_flush_with_return_true_when_active
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start();
echo 'z';
echo ob_end_flush() ? 'ok' : 'bad';
echo ob_get_level();
