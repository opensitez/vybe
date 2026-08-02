<?php
// vybe-test: php/php_functions_arrow_first_class_callables/test_php_named_arguments_array_unpacking
// origin: languages/php/tests/php/test_php_functions_arrow_first_class_callables.rs
// vybe-test-mode: compile

function configure(string $host, int $port = 8080, bool $ssl = false) {
    return "$host:$port ssl=" . ($ssl ? "yes" : "no");
}

$params = ["ssl" => true, "host" => "localhost"];
echo configure(...$params);
