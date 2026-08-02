<?php
// vybe-test: php/php_oop_constructor_promotion_readonly/test_php_constructor_promotion_default_expressions
// origin: languages/php/tests/php/test_php_oop_constructor_promotion_readonly.rs
// vybe-test-mode: compile

class Config {
    public function __construct(
        public array $options = ["debug" => true],
        public string $env = "production"
    ) {}
}

$c = new Config();
echo $c->env;
