<?php
// vybe-test: php/output_buffering/ob_start_without_callback_default_gets_buffer
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start();
echo 'abc';
echo ob_get_clean();
