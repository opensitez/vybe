<?php
// vybe-test: php/filters/filter_has_var_env
// origin: languages/php/tests/php/test_filters.rs
// vybe-test-mode: compile

// INPUT_ENV checks environment
$has = filter_has_var(INPUT_ENV, 'PATH');
echo is_bool($has) ? 'bool result' : 'non-bool';
