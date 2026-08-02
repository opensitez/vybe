<?php
// vybe-test: php/output_buffering/ob_get_length_with_multi_byte
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start();
echo "éclair";
$len = ob_get_length();
ob_end_clean();
echo $len . '|' . strlen("éclair");
