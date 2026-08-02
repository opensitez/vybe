<?php
// vybe-test: php/oop/enum_method
// origin: languages/php/tests/php/test_oop.rs
// vybe-test-mode: compile

enum Status {
    case Active;
    case Inactive;
    public function label() { return $this->name; }
}
echo Status::Active->label();
