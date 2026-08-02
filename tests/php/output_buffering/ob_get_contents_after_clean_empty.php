<?php
// vybe-test: php/output_buffering/ob_get_contents_after_clean_empty
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start();
echo 'temp';
ob_clean();
$r = ob_get_contents() === '' ? 'clean' : 'dirty';
ob_end_clean();
echo $r;
