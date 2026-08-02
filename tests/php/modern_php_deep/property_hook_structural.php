<?php
// vybe-test: php/modern_php_deep/property_hook_structural
// origin: languages/php/tests/php/test_modern_php_deep.rs
// vybe-test-mode: compile

class Temperature {
    public float $celsius {
        get { return $this->celsius; }
        set(float $value) { $this->celsius = $value; }
    }
    public float $fahrenheit {
        get { return $this->celsius * 9/5 + 32; }
    }
}
$t = new Temperature();
$t->celsius = 100.0;
echo $t->fahrenheit;
