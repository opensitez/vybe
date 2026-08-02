<?php
// vybe-test: php/php_functions_arrow_fn_variadic_named/test_php_named_arguments_in_constructor_call
// origin: languages/php/tests/php/test_php_functions_arrow_fn_variadic_named.rs
// vybe-test-mode: compile

class ServerConfig {
    public function __construct(
        public string $host,
        public int $port = 80,
        public int $timeout = 30
    ) {}
}

$config = new ServerConfig(timeout: 60, host: "127.0.0.1");
echo "{$config->host}:{$config->port} t={$config->timeout}";
