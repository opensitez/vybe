<?php
// vybe-test: php/php_oop_final_readonly_property_promotion/test_php_asymmetric_visibility_constructor_promotion_php84
// origin: languages/php/tests/php/test_php_oop_final_readonly_property_promotion.rs
// vybe-test-mode: compile

class CounterService {
    public function __construct(
        public private(set) int $count = 0
    ) {}

    public function increment(): void {
        $this->count++;
    }
}

$cs = new CounterService();
$cs->increment();
echo $cs->count;
