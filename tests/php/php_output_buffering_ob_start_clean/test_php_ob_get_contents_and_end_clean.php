<?php
// vybe-test: php/php_output_buffering_ob_start_clean/test_php_ob_get_contents_and_end_clean
// origin: languages/php/tests/php/test_php_output_buffering_ob_start_clean.rs

ob_start();
echo "Internal output";
$contents = ob_get_contents();
ob_end_clean();
echo "Retrieved: $contents";
