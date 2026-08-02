<?php
// vybe-test: php/php_functions_arrow_first_class_callables/test_php_variadic_parameter_with_type_hint
// origin: languages/php/tests/php/test_php_functions_arrow_first_class_callables.rs
// vybe-test-mode: compile

function concatenate(string $delim, string ...$words): string {
    return implode($delim, $words);
}

echo concatenate("-", "a", "b", "c");
