<?php
// vybe-test: php/php_enums_backed_tryfrom_cases/test_php81_enum_constant_access
// origin: languages/php/tests/php/test_php_enums_backed_tryfrom_cases.rs
// vybe-test-mode: compile

enum Feature: string {
    case Beta = "beta";
    case Stable = "stable";

    public const DEFAULT_FEATURE = self::Stable;
}

echo Feature::DEFAULT_FEATURE->value;
