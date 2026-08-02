<?php
// vybe-test: php/php_oop_constructor_promotion_readonly/test_php_asymmetric_property_visibility_php84
// origin: languages/php/tests/php/test_php_oop_constructor_promotion_readonly.rs
// vybe-test-mode: compile

class Account {
    public private(set) float $balance = 0.0;

    public function deposit(float $amount): void {
        $this->balance += $amount;
    }
}

$acc = new Account();
$acc->deposit(100.0);
echo $acc->balance;
