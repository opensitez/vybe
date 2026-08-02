<?php
// vybe-test: php/output_buffering/ob_get_status_without_level_key_returns_map
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start();
$status = ob_get_status(false);
ob_end_clean();
echo is_array($status) ? (array_key_exists('level', $status) ? 'ok' : 'nolvl') : 'bad';
