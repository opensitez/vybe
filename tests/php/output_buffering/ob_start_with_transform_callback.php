<?php
// vybe-test: php/output_buffering/ob_start_with_transform_callback
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start(fn(string $buf): string => strtoupper($buf));
echo 'hello';
ob_end_flush();
