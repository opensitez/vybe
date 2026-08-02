<?php
// vybe-test: php/output_buffering/ob_get_flush_with_active_buffer_returns_content
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start();
echo 'xy';
echo ob_get_flush();
