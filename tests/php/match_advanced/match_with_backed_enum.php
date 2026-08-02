<?php
// vybe-test: php/match_advanced/match_with_backed_enum
// origin: languages/php/tests/php/test_match_advanced.rs
// vybe-test-mode: compile

enum Status: string { case Active = 'A'; case Inactive = 'I'; case Pending = 'P'; }
$s = Status::Active;
$label = match($s) {
    Status::Active   => 'Active',
    Status::Inactive => 'Inactive',
    Status::Pending  => 'Pending',
};
echo $label;
