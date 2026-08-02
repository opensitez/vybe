<?php
// vybe-test: php/output_buffering/ob_callback_can_return_empty_string
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start(function(string $buf): string { return ''; });
echo 'drop';
echo ob_get_clean();
