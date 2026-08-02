<?php
// vybe-test: php/output_buffering/ob_ob_get_clean_discards_outer_when_nested
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start();
echo 'outer';
ob_start();
echo 'inner';
ob_end_clean();
$value = ob_get_clean();
echo $value;
