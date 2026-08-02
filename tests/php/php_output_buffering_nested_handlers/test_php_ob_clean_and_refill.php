<?php
// vybe-test: php/php_output_buffering_nested_handlers/test_php_ob_clean_and_refill
// origin: languages/php/tests/php/test_php_output_buffering_nested_handlers.rs
// vybe-test-mode: compile

ob_start();
echo "wrong data";
ob_clean();
echo "correct data";
$res = ob_get_clean();
echo $res;
