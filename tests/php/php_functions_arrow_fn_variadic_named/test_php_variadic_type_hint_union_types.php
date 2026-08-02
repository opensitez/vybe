<?php
// vybe-test: php/php_functions_arrow_fn_variadic_named/test_php_variadic_type_hint_union_types
// origin: languages/php/tests/php/test_php_functions_arrow_fn_variadic_named.rs
// vybe-test-mode: compile

function stringify(int|float ...$values): array {
    return array_map(fn($v) => (string)$v, $values);
}

$res = stringify(1, 2.5, 3);
echo implode("-", $res);
