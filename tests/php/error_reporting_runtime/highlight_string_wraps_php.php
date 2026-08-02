<?php
// vybe-test: php/error_reporting_runtime/highlight_string_wraps_php
// origin: languages/php/tests/php/test_error_reporting_runtime.rs

ob_start();
highlight_string('<?php echo 1;');
$out = ob_get_clean();
echo str_contains($out, 'php') ? 'html' : 'no';
