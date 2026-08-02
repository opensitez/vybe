<?php
// vybe-test: php/php84/property_hook_set_validation
// origin: languages/php/tests/php/test_php84.rs
// vybe-test-mode: compile

class User {
    public string $email {
        set(string $value) {
            if (!str_contains($value, '@')) {
                throw new \InvalidArgumentException("Invalid email: $value");
            }
            $this->email = strtolower($value);
        }
    }
}
$u = new User();
$u->email = 'Alice@Example.COM';
echo $u->email;
