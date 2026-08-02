<?php
// vybe-test: php/function_builtins/call_user_func_by_name
// origin: languages/php/tests/php/test_function_builtins.rs
// vybe-test-mode: compile

function greet(string $name): string {
    return 'Hello, ' . $name;
}
echo call_user_func('greet', 'World');
