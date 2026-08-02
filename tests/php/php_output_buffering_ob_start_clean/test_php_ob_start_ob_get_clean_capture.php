<?php
// vybe-test: php/php_output_buffering_ob_start_clean/test_php_ob_start_ob_get_clean_capture
// origin: languages/php/tests/php/test_php_output_buffering_ob_start_clean.rs

ob_start();
echo "Buffered HTML Content";
$captured = ob_get_clean();
echo "Captured: $captured";
