<?php
// vybe-test: php/output_functions/vprintf_with_array
// origin: languages/php/tests/php/test_output_functions.rs
// vybe-test-mode: compile

$args = ['Charlie', 7, 99.9];
$written = vprintf("Player: %s, Level: %d, HP: %.1f\n", $args);
echo $written > 0 ? 'ok' : 'fail';
