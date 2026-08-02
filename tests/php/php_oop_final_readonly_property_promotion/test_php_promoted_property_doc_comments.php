<?php
// vybe-test: php/php_oop_final_readonly_property_promotion/test_php_promoted_property_doc_comments
// origin: languages/php/tests/php/test_php_oop_final_readonly_property_promotion.rs
// vybe-test-mode: compile

class Customer {
    public function __construct(
        /** @var string Customer full name */
        public string $name,
        /** @var string Customer email address */
        public string $email
    ) {}
}

$c = new Customer("Bob", "bob@example.com");
echo "$c->name <$c->email>";
