<?php
// vybe-test: php/php_enums_backed_methods_attributes/test_php81_enum_cases_array
// origin: languages/php/tests/php/test_php_enums_backed_methods_attributes.rs
// vybe-test-mode: compile

enum Direction {
    case North;
    case South;
    case East;
    case West;
}

$cases = Direction::cases();
foreach ($cases as $case) {
    echo $case->name . "\n";
}
