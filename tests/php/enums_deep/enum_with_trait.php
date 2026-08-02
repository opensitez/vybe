<?php
// vybe-test: php/enums_deep/enum_with_trait
// origin: languages/php/tests/php/test_enums_deep.rs
// vybe-test-mode: compile

trait HasLabel {
    public function label(): string {
        return ucfirst(strtolower($this->name));
    }
}
enum Status {
    use HasLabel;
    case Active;
    case Inactive;
    case Pending;
}
echo Status::Active->label();
echo Status::Pending->label();
