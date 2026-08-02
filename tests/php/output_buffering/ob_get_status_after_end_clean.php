<?php
// vybe-test: php/output_buffering/ob_get_status_after_end_clean
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start();
echo 'a';
ob_end_clean();
$status = ob_get_status(false);
echo is_array($status) ? 'arr' : 'bad';
echo '|';
echo $status['level'];
