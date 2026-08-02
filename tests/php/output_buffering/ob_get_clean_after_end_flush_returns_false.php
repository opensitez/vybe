<?php
// vybe-test: php/output_buffering/ob_get_clean_after_end_flush_returns_false
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start();
echo 'flush-me';
ob_end_flush();
echo ob_get_clean() === false ? 'closed' : 'open';
