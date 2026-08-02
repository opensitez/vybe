<?php
// vybe-test: php/type_functions_extended/is_infinite_check
// origin: languages/php/tests/php/test_type_functions_extended.rs
// vybe-test-mode: compile

echo is_infinite(INF) ? 'yes' : 'no';
echo is_infinite(3.14) ? 'yes' : 'no';
