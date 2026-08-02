<?php
// vybe-test: php/php_traits_composition_conflict_resolution/test_php_class_method_precedence_over_trait
// origin: languages/php/tests/php/test_php_traits_composition_conflict_resolution.rs
// vybe-test-mode: compile

trait Greeting {
    public function hello() { return "Trait Hello"; }
}

class CustomGreeting {
    use Greeting;
    public function hello() { return "Class Hello"; }
}

$cg = new CustomGreeting();
echo $cg->hello(); // Class method overrides trait method!
