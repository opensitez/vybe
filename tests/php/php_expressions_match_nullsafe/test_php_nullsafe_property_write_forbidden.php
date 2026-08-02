<?php
// vybe-test: php/php_expressions_match_nullsafe/test_php_nullsafe_property_write_forbidden
// origin: languages/php/tests/php/test_php_expressions_match_nullsafe.rs
// vybe-test-mode: compile

class Container {
    public ?stdClass $inner = null;
}

$c = new Container();
    // Nullsafe operator cannot be used on left hand side of assignment
$val = $c?->inner?->prop;
