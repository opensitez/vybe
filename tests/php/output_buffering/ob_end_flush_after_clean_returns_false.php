<?php
// vybe-test: php/output_buffering/ob_end_flush_after_clean_returns_false
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start();
echo 'x';
ob_end_clean();
echo ob_end_flush() ? 'closed' : 'false';
