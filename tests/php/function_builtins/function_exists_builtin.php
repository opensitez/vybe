<?php
// vybe-test: php/function_builtins/function_exists_builtin
// origin: languages/php/tests/php/test_function_builtins.rs
// vybe-test-mode: compile

echo function_exists('strlen') ? 'yes' : 'no';
echo function_exists('array_map') ? 'yes' : 'no';
echo function_exists('no_such_function_xyz') ? 'yes' : 'no';
