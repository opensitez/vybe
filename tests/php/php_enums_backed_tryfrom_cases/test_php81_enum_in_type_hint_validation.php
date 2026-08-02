<?php
// vybe-test: php/php_enums_backed_tryfrom_cases/test_php81_enum_in_type_hint_validation
// origin: languages/php/tests/php/test_php_enums_backed_tryfrom_cases.rs
// vybe-test-mode: compile

enum Environment: string { case Dev = "dev"; case Prod = "prod"; }

function setEnv(Environment $env): string {
    return "Setting: " . $env->value;
}

echo setEnv(Environment::Dev);
