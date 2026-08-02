<?php
// vybe-test: php/output_buffering/ob_flush_without_content_returns_empty_string
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start();
echo ob_get_length();
echo '|';
ob_flush();
echo ob_get_length();
ob_end_clean();
