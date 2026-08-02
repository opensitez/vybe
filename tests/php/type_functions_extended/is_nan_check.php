<?php
// vybe-test: php/type_functions_extended/is_nan_check
// origin: languages/php/tests/php/test_type_functions_extended.rs
// vybe-test-mode: compile

echo is_nan(NAN) ? 'yes' : 'no';
echo is_nan(0.0) ? 'yes' : 'no';
