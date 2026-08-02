<?php
// vybe-test: php/php_expressions_match_nullsafe/test_php_match_expression_boolean_expressions
// origin: languages/php/tests/php/test_php_expressions_match_nullsafe.rs
// vybe-test-mode: compile

$age = 25;
$category = match (true) {
    $age < 13 => "child",
    $age < 20 => "teen",
    $age >= 20 => "adult",
};
echo $category;
