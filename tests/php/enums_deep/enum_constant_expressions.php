<?php
// vybe-test: php/enums_deep/enum_constant_expressions
// origin: languages/php/tests/php/test_enums_deep.rs
// vybe-test-mode: compile

enum Permission: int {
    case Read    = 1;
    case Write   = 2;
    case Execute = 4;
    const int ALL = self::Read->value | self::Write->value | self::Execute->value;
}
echo Permission::ALL;
