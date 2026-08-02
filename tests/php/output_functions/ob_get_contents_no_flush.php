<?php
// vybe-test: php/output_functions/ob_get_contents_no_flush
// origin: languages/php/tests/php/test_output_functions.rs
// vybe-test-mode: compile

ob_start();
echo 'peek';
$buf = ob_get_contents();
ob_end_clean();
echo 'peeked: ' . $buf;
