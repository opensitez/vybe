<?php
// vybe-test: php/output_buffering/ob_callback_drops_output_when_empty
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start(fn(string $buf): string => '');
echo 'abc';
ob_end_clean();
echo 'after';
