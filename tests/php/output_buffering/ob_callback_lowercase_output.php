<?php
// vybe-test: php/output_buffering/ob_callback_lowercase_output
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start(fn(string $buf): string => strtoupper(strtolower($buf)));
echo 'MiXeD';
ob_end_flush();
