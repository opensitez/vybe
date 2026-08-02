<?php
// vybe-test: php/output_buffering/ob_start_with_callable_function_name
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start('str_rot13');
echo 'uryyb';
ob_end_flush();
