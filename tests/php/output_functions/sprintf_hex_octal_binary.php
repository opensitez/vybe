<?php
// vybe-test: php/output_functions/sprintf_hex_octal_binary
// origin: languages/php/tests/php/test_output_functions.rs
// vybe-test-mode: compile

echo sprintf('%x', 255);
echo sprintf('%X', 255);
echo sprintf('%o', 8);
echo sprintf('%b', 10);
echo sprintf('%08b', 10);
