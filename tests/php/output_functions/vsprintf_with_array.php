<?php
// vybe-test: php/output_functions/vsprintf_with_array
// origin: languages/php/tests/php/test_output_functions.rs
// vybe-test-mode: compile

$args = ['PHP', '8.3', 'Stable'];
$result = vsprintf('%s %s (%s)', $args);
echo $result;
