<?php
// vybe-test: php/output_buffering/ob_clean_keeps_buffer_handle_open
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start();
echo 'abc';
ob_clean();
$lvl = ob_get_level();
ob_get_clean();
echo $lvl . '|';
