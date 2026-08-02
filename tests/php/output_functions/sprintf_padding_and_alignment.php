<?php
// vybe-test: php/output_functions/sprintf_padding_and_alignment
// origin: languages/php/tests/php/test_output_functions.rs
// vybe-test-mode: compile

$left  = sprintf('%-10s|', 'hi');
$right = sprintf('%010d',  42);
echo $left;
echo $right;
