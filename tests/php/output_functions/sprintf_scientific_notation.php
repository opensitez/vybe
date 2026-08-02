<?php
// vybe-test: php/output_functions/sprintf_scientific_notation
// origin: languages/php/tests/php/test_output_functions.rs
// vybe-test-mode: compile

echo sprintf('%e', 123456.789);
echo sprintf('%E', 0.000123);
echo sprintf('%.3e', 9876.5432);
