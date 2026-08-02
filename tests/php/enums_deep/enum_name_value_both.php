<?php
// vybe-test: php/enums_deep/enum_name_value_both
// origin: languages/php/tests/php/test_enums_deep.rs
// vybe-test-mode: compile

enum Currency: string {
    case USD = 'US Dollar';
    case EUR = 'Euro';
    case GBP = 'British Pound';
}
foreach (Currency::cases() as $c) {
    echo "{$c->name}: {$c->value}\n";
}
