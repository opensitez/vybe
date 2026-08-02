<?php
// vybe-test: php/type_checks_extra/is_iterable_array_check
// origin: languages/php/tests/php/test_type_checks_extra.rs
// vybe-test-mode: compile

echo is_iterable([1, 2, 3]) ? 'yes' : 'no';
echo is_iterable(42) ? 'yes' : 'no';
