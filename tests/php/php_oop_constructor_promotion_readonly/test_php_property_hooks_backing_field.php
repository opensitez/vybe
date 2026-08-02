<?php
// vybe-test: php/php_oop_constructor_promotion_readonly/test_php_property_hooks_backing_field
// origin: languages/php/tests/php/test_php_oop_constructor_promotion_readonly.rs
// vybe-test-mode: compile

class Temperature {
    public float $celsius {
        set {
            if ($value < -273.15) {
                throw new InvalidArgumentException("Below absolute zero");
            }
            $this->celsius = $value;
        }
    }
}

$t = new Temperature();
$t->celsius = 25.0;
echo $t->celsius;
