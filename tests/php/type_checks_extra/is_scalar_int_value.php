<?php
// vybe-test: php/type_checks_extra/is_scalar_int_value
// origin: languages/php/tests/php/test_type_checks_extra.rs
// vybe-test-mode: compile

echo is_scalar(42) ? 'yes' : 'no';
echo is_scalar([1, 2]) ? 'yes' : 'no';
