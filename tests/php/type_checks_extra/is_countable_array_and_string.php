<?php
// vybe-test: php/type_checks_extra/is_countable_array_and_string
// origin: languages/php/tests/php/test_type_checks_extra.rs
// vybe-test-mode: compile

echo is_countable([1, 2, 3]) ? 'yes' : 'no';
echo is_countable('hello') ? 'yes' : 'no';
