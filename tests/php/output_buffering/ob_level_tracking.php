<?php
// vybe-test: php/output_buffering/ob_level_tracking
// origin: languages/php/tests/php/test_output_buffering.rs
// vybe-test-mode: compile

$levels = [];
$levels[] = ob_get_level();
ob_start();
$levels[] = ob_get_level();
ob_start();
$levels[] = ob_get_level();
ob_end_clean();
ob_end_clean();
$levels[] = ob_get_level();
echo implode(',', $levels);
