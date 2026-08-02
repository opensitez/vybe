<?php
// vybe-test: php/output_buffering/ob_end_clean_returns_false_after_close
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start();
ob_end_clean();
echo ob_end_clean() ? 'open' : 'closed';
