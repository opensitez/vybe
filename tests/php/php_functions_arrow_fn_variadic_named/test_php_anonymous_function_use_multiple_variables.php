<?php
// vybe-test: php/php_functions_arrow_fn_variadic_named/test_php_anonymous_function_use_multiple_variables
// origin: languages/php/tests/php/test_php_functions_arrow_fn_variadic_named.rs
// vybe-test-mode: compile

$prefix = "LOG";
$suffix = "END";

$log = function(string $msg) use ($prefix, $suffix) {
    return "$prefix: $msg ($suffix)";
};

echo $log("Message body");
