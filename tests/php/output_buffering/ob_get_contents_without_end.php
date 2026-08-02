<?php
// vybe-test: php/output_buffering/ob_get_contents_without_end
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start();
echo 'stay';
$s = ob_get_contents();
ob_end_clean();
echo $s;
