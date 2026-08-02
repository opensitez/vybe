<?php
// vybe-test: php/php_enums_backed_methods_attributes/test_php81_enum_in_match_expression
// origin: languages/php/tests/php/test_php_enums_backed_methods_attributes.rs
// vybe-test-mode: compile

enum State { case Draft; case Published; case Archived; }

$state = State::Published;
$label = match($state) {
    State::Draft => "Draft Document",
    State::Published => "Live Document",
    State::Archived => "Archived",
};
echo $label;
