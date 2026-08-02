<?php
// vybe-test: php/php84/property_hook_with_backing
// origin: languages/php/tests/php/test_php84.rs
// vybe-test-mode: compile

class Product {
    private float $_price = 0.0;
    public float $price {
        get { return $this->_price; }
        set(float $value) {
            if ($value < 0) throw new \RangeException("Price cannot be negative");
            $this->_price = round($value, 2);
        }
    }
}
$p = new Product();
$p->price = 9.999;
echo $p->price;
