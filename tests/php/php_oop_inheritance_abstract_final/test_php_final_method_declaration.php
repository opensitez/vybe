<?php
// vybe-test: php/php_oop_inheritance_abstract_final/test_php_final_method_declaration
// origin: languages/php/tests/php/test_php_oop_inheritance_abstract_final.rs
// vybe-test-mode: compile

class AuthManager {
    final public function hashPassword(string $pwd): string {
        return md5($pwd);
    }
}

class CustomAuth extends AuthManager {
    public function login(): void {}
}

$ca = new CustomAuth();
echo $ca->hashPassword("12345");
