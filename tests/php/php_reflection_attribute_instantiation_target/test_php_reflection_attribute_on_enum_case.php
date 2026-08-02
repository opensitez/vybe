<?php
// vybe-test: php/php_reflection_attribute_instantiation_target/test_php_reflection_attribute_on_enum_case
// origin: languages/php/tests/php/test_php_reflection_attribute_instantiation_target.rs
// vybe-test-mode: compile

#[Attribute(Attribute::TARGET_CLASS_CONSTANT)]
class Label { public function __construct(public string $text) {} }

enum Status {
    #[Label("Pending Approval")]
    case Pending;
}

$re = new ReflectionEnum(Status::class);
$case = $re->getCase("Pending");
$attrs = $case->getAttributes(Label::class);
echo count($attrs);
