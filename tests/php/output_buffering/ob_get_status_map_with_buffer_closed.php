<?php
// vybe-test: php/output_buffering/ob_get_status_map_with_buffer_closed
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start();
echo 'x';
ob_get_status();
ob_end_clean();
$status = ob_get_status();
echo is_array($status) && array_key_exists('level', $status) ? 'arr' : 'not';
echo '|' . $status['level'];
