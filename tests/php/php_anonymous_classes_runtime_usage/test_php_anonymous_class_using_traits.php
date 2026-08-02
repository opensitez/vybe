<?php
// vybe-test: php/php_anonymous_classes_runtime_usage/test_php_anonymous_class_using_traits
// origin: languages/php/tests/php/test_php_anonymous_classes_runtime_usage.rs
// vybe-test-mode: compile

trait IdentityTrait {
    public function getId(): int { return 999; }
}

$entity = new class {
    use IdentityTrait;
};

echo $entity->getId();
