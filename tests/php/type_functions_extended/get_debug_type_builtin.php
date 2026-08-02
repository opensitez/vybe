<?php
// vybe-test: php/type_functions_extended/get_debug_type_builtin
// origin: languages/php/tests/php/test_type_functions_extended.rs
// vybe-test-mode: compile

echo get_debug_type(42);
echo get_debug_type(3.14);
echo get_debug_type('hello');
echo get_debug_type(null);
