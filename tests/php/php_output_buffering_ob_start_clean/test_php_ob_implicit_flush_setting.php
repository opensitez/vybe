<?php
// vybe-test: php/php_output_buffering_ob_start_clean/test_php_ob_implicit_flush_setting
// origin: languages/php/tests/php/test_php_output_buffering_ob_start_clean.rs
// vybe-test-mode: compile

ob_implicit_flush(true);
echo "immediate output";
ob_implicit_flush(false);
