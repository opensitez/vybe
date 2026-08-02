<?php
// vybe-test: php/php_ob_gzhandler_compression/test_php_ob_get_length_returns_buffer_byte_count
// origin: languages/php/tests/php/test_php_ob_gzhandler_compression.rs
// vybe-test-mode: compile

ob_start();
echo "1234567890";
$len = ob_get_length();
ob_end_clean();
echo $len === 10 ? "BUFFER_LEN_10_OK" : "FAIL";
