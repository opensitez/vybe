<?php
// vybe-test: php/php_exception_types_spl_hierarchy/test_php_domain_exception_business_logic_violation
// origin: languages/php/tests/php/test_php_exception_types_spl_hierarchy.rs
// vybe-test-mode: compile

function calculateDiscount(float $price, float $discount) {
    if ($discount > $price) {
        throw new DomainException("Discount exceeds total price");
    }
    return $price - $discount;
}

try {
    calculateDiscount(10.0, 15.0);
} catch (DomainException $e) {
    echo $e->getMessage();
}
