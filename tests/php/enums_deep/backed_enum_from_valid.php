<?php
// vybe-test: php/enums_deep/backed_enum_from_valid
// origin: languages/php/tests/php/test_enums_deep.rs
// vybe-test-mode: compile

enum Status: string {
    case Active   = 'active';
    case Inactive = 'inactive';
    case Banned   = 'banned';
}
$s = Status::from('active');
echo $s->name . ':' . $s->value;
