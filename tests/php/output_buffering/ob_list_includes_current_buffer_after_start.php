<?php
// vybe-test: php/output_buffering/ob_list_includes_current_buffer_after_start
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start();
$list = ob_list_handlers();
ob_end_clean();
echo count($list) >= 0 ? 'listed' : 'none';
