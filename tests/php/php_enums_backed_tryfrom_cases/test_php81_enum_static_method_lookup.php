<?php
// vybe-test: php/php_enums_backed_tryfrom_cases/test_php81_enum_static_method_lookup
// origin: languages/php/tests/php/test_php_enums_backed_tryfrom_cases.rs
// vybe-test-mode: compile

enum Severity: int {
    case Low = 1;
    case Medium = 2;
    case High = 3;

    public static function fromName(string $name): ?self {
        foreach (self::cases() as $case) {
            if ($case->name === $name) return $case;
        }
        return null;
    }
}

$s = Severity::fromName("High");
echo $s ? $s->value : 0;
