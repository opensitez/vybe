<?php
// vybe-test: php/php_traits_composition_conflict_resolution/test_php_trait_parent_class_precedence
// origin: languages/php/tests/php/test_php_traits_composition_conflict_resolution.rs
// vybe-test-mode: compile

class BaseClass {
    public function say() { return "Base"; }
}

trait TraitSay {
    public function say() { return "Trait"; }
}

class ChildClass extends BaseClass {
    use TraitSay;
}

$c = new ChildClass();
echo $c->say(); // Trait method overrides parent class method!
