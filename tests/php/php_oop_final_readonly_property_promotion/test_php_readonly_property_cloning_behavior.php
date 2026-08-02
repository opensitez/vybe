<?php
// vybe-test: php/php_oop_final_readonly_property_promotion/test_php_readonly_property_cloning_behavior
// origin: languages/php/tests/php/test_php_oop_final_readonly_property_promotion.rs
// vybe-test-mode: compile

class Order {
    public readonly DateTimeImmutable $createdAt;
    public function __construct() {
        $this->createdAt = new DateTimeImmutable();
    }
    public function __clone() {
        // Readonly properties can be modified during __clone in PHP 8.3+
    }
}

$o1 = new Order();
$o2 = clone $o1;
echo get_class($o2);
