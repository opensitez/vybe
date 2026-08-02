<?php
// vybe-test: php/error_handling_deep/at_operator_suppress
// origin: languages/php/tests/php/test_error_handling_deep.rs
// vybe-test-mode: compile

// @ suppresses errors from the expression
$result = @file_get_contents('/nonexistent/file/path');
echo $result === false ? 'failed silently' : 'unexpected success';
