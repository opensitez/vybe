<?php
// vybe-test: php/php_enums_backed_methods_attributes/test_php81_enum_static_methods
// origin: languages/php/tests/php/test_php_enums_backed_methods_attributes.rs
// vybe-test-mode: compile

enum Role: string {
    case Admin = "admin";
    case User = "user";

    public static function values(): array {
        return array_column(self::cases(), "value");
    }
}

print_r(Role::values());
