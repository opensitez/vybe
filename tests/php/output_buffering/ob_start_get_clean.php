<?php
// vybe-test: php/output_buffering/ob_start_get_clean
// origin: languages/php/tests/php/test_output_buffering.rs
// vybe-test-mode: compile

ob_start();
echo "buffered content";
$content = ob_get_clean();
echo "captured: $content";
