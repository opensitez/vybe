<?php
// vybe-test: php/php_oop_constructor_promotion_readonly/test_php_constructor_promotion_with_attributes
// origin: languages/php/tests/php/test_php_oop_constructor_promotion_readonly.rs
// vybe-test-mode: compile

#[Attribute]
class Validate {
    public function __construct(public string $rule) {}
}

class Product {
    public function __construct(
        #[Validate("min:1")]
        public string $title,
        #[Validate("gt:0")]
        public float $price
    ) {}
}

$p = new Product("Widget", 19.99);
echo $p->title;
