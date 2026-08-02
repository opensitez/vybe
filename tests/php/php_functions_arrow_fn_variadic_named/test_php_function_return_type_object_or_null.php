<?php
// vybe-test: php/php_functions_arrow_fn_variadic_named/test_php_function_return_type_object_or_null
// origin: languages/php/tests/php/test_php_functions_arrow_fn_variadic_named.rs
// vybe-test-mode: compile

function findService(string $name): ?object {
    if ($name === "db") return new stdClass();
    return null;
}

echo is_object(findService("db")) ? "FOUND" : "NOT_FOUND";
