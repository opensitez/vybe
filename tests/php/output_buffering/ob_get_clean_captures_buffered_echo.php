<?php
// vybe-test: php/output_buffering/ob_get_clean_captures_buffered_echo
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start();
echo 'buf';
$c = ob_get_clean();
echo $c;
