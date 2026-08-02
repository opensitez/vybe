<?php
// vybe-test: php/output_buffering/ob_get_status_depth_false
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start();
ob_start();
$lvl = ob_get_status()['level'];
ob_end_clean();
ob_end_clean();
echo $lvl;
