<?php
// vybe-test: php/output_buffering/ob_end_clean_nested_discard_outer_buffer
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start();
echo 'outer';
ob_start();
echo 'inner';
ob_end_clean();
echo '|';
$o = ob_get_clean();
echo $o;
