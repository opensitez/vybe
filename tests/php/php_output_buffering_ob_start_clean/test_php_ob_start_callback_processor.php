<?php
// vybe-test: php/php_output_buffering_ob_start_clean/test_php_ob_start_callback_processor
// origin: languages/php/tests/php/test_php_output_buffering_ob_start_clean.rs

ob_start(function($buffer) {
    return strtoupper($buffer);
});
echo "lowercase text";
ob_end_flush();
