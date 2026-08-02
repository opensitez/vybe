<?php
// vybe-test: php/php_oop_final_readonly_property_promotion/test_php_promoted_property_callable_default_error
// origin: languages/php/tests/php/test_php_oop_final_readonly_property_promotion.rs
// vybe-test-mode: compile

class Task {
    public function __construct(
        public string $title,
        public int $priority = 1
    ) {}
}

$t = new Task("Write Tests");
echo $t->title;
