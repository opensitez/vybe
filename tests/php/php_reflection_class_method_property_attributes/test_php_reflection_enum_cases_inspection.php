<?php
// vybe-test: php/php_reflection_class_method_property_attributes/test_php_reflection_enum_cases_inspection
// origin: languages/php/tests/php/test_php_reflection_class_method_property_attributes.rs
// vybe-test-mode: compile

enum Suit: string {
    case Hearts = "H";
    case Diamonds = "D";
}

$re = new ReflectionEnum(Suit::class);
echo $re->isBacked() ? "BACKED" : "PURE";
$cases = $re->getCases();
foreach ($cases as $case) {
    echo $case->getName() . "=" . $case->getValue()->value . "\n";
}
