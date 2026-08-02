<?php
// vybe-test: php/output_buffering/ob_list_handlers_length_tracks_start
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start();
ob_start(function($buf) { return strtoupper($buf); });
$handlers = ob_list_handlers();
ob_end_clean();
ob_end_clean();
echo is_array($handlers) ? count($handlers) : 0;
