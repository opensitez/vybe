<?php
// vybe-test: php/output_buffering/ob_get_contents
// origin: languages/php/tests/php/test_output_buffering.rs
// vybe-test-mode: compile

ob_start();
echo "hello";
echo " world";
$buf = ob_get_contents();
ob_end_clean();
echo strlen($buf);
