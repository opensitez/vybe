<?php
// vybe-test: php/php_ob_gzhandler_compression/test_php_ob_start_flags_cleanable_flushable
// origin: languages/php/tests/php/test_php_ob_gzhandler_compression.rs
// vybe-test-mode: compile

ob_start(null, 0, PHP_OUTPUT_HANDLER_CLEANABLE | PHP_OUTPUT_HANDLER_FLUSHABLE);
echo "Flushes";
$cleared = ob_end_clean();
echo $cleared ? "CLEANABLE_FLAG_OK" : "FAIL";
