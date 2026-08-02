<?php
// vybe-test: php/php_enums_backed_methods_attributes/test_php81_enum_constant_expressions
// origin: languages/php/tests/php/test_php_enums_backed_methods_attributes.rs
// vybe-test-mode: compile

enum Size {
    case Small;
    case Medium;
    case Large;

    public const Default = self::Medium;
}

echo Size::Default->name;
