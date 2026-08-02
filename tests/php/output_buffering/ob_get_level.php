<?php
// vybe-test: php/output_buffering/ob_get_level
// origin: languages/php/tests/php/test_output_buffering.rs
// vybe-test-mode: compile

echo ob_get_level();
ob_start();
echo ob_get_level();
ob_end_clean();
echo ob_get_level();
