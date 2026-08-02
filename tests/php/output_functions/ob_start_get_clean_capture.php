<?php
// vybe-test: php/output_functions/ob_start_get_clean_capture
// origin: languages/php/tests/php/test_output_functions.rs
// vybe-test-mode: compile

ob_start();
echo 'captured output';
$content = ob_get_clean();
echo 'got: ' . $content;
