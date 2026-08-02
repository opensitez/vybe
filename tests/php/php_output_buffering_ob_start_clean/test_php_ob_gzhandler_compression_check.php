<?php
// vybe-test: php/php_output_buffering_ob_start_clean/test_php_ob_gzhandler_compression_check
// origin: languages/php/tests/php/test_php_output_buffering_ob_start_clean.rs
// vybe-test-mode: compile

if (extension_loaded('zlib')) {
    ob_start('ob_gzhandler');
    echo "Compressed page content";
    ob_end_clean();
}
