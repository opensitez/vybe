<?php
// vybe-test: php/output_buffering/ob_flush_keeps_buffer
// origin: languages/php/tests/php/test_output_buffering.rs
// vybe-test-mode: compile

ob_start();
echo "part one";
ob_flush();
echo "part two";
$all = ob_get_clean();
echo $all;
