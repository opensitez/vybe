<?php
// vybe-test: php/enums_deep/enum_trait_with_property_check
// origin: languages/php/tests/php/test_enums_deep.rs
// vybe-test-mode: compile

trait HasPriority {
    public function isHighPriority(): bool {
        return match($this) {
            self::Critical, self::High => true,
            default => false,
        };
    }
}
enum Severity {
    use HasPriority;
    case Critical;
    case High;
    case Medium;
    case Low;
}
echo Severity::Critical->isHighPriority() ? 'high' : 'low';
echo Severity::Low->isHighPriority() ? 'high' : 'low';
