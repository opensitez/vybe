<?php
// vybe-test: php/php_enums_backed_tryfrom_cases/test_php81_enum_in_switch_case
// origin: languages/php/tests/php/test_php_enums_backed_tryfrom_cases.rs
// vybe-test-mode: compile

enum Mode { case Read; case Write; case Admin; }

$m = Mode::Write;
$desc = "";
switch ($m) {
    case Mode::Read: $desc = "Read Only"; break;
    case Mode::Write: $desc = "Read Write"; break;
    case Mode::Admin: $desc = "Full Access"; break;
}
echo $desc;
