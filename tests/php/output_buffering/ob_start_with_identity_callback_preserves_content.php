<?php
// vybe-test: php/output_buffering/ob_start_with_identity_callback_preserves_content
// origin: languages/php/tests/php/test_output_buffering.rs

ob_start(fn(string $buf): string => $buf);
echo 'z';
echo ob_get_clean();
