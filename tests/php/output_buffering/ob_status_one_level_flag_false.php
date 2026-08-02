<?php
// vybe-test: php/output_buffering/ob_status_one_level_flag_false
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start();
$status = ob_get_status();
ob_end_clean();
echo is_array($status) ? 'arr' : 'no';
