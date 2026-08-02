<?php
// vybe-test: php/output_buffering/ob_callback_transform
// origin: languages/php/tests/php/test_output_buffering.rs
// vybe-test-mode: compile

ob_start(function(string $buf): string {
    return str_replace(' ', '_', $buf);
});
echo "hello world again";
ob_end_flush();
