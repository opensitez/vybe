<?php
// vybe-test: php/output_buffering/ob_implicit_flush_true_does_not_disable_ob_start
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start();
ob_implicit_flush(true);
echo 'z';
$lvl = ob_get_level();
ob_end_clean();
echo $lvl >= 1 ? 'active' : 'off';
