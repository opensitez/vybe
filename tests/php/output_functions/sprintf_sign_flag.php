<?php
// vybe-test: php/output_functions/sprintf_sign_flag
// origin: languages/php/tests/php/test_output_functions.rs
// vybe-test-mode: compile

echo sprintf('%+d', 42);
echo sprintf('%+d', -42);
echo sprintf('%+.2f', 3.14);
echo sprintf('%+.2f', -3.14);
