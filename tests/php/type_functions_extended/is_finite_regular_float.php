<?php
// vybe-test: php/type_functions_extended/is_finite_regular_float
// origin: languages/php/tests/php/test_type_functions_extended.rs
// vybe-test-mode: compile

echo is_finite(3.14) ? 'yes' : 'no';
echo is_finite(INF) ? 'yes' : 'no';
