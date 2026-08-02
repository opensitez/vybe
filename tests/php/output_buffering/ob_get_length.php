<?php
// vybe-test: php/output_buffering/ob_get_length
// origin: languages/php/tests/php/test_output_buffering.rs
// vybe-test-mode: compile

ob_start();
echo "twelve chars";
echo ob_get_length();
ob_end_clean();
