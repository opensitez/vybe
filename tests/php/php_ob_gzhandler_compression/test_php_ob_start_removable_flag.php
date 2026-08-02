<?php
// vybe-test: php/php_ob_gzhandler_compression/test_php_ob_start_removable_flag
// origin: languages/php/tests/php/test_php_ob_gzhandler_compression.rs
// vybe-test-mode: compile

ob_start(null, 0, PHP_OUTPUT_HANDLER_REMOVABLE);
echo "Removable";
$ended = ob_end_clean();
echo $ended ? "REMOVABLE_FLAG_OK" : "FAIL";
