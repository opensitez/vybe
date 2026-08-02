<?php
// vybe-test: php/php_oop_property_hooks_asymmetric_visibility/test_php84_asymmetric_visibility_readonly_combination
// origin: languages/php/tests/php/test_php_oop_property_hooks_asymmetric_visibility.rs
// vybe-test-mode: compile

class Token {
    public private(set) readonly string $hash;

    public function __construct(string $secret) {
        $this->hash = md5($secret);
    }
}

$t = new Token("my_secret");
echo $t->hash;
