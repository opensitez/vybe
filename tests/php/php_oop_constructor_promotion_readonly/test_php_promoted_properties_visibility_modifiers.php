<?php
// vybe-test: php/php_oop_constructor_promotion_readonly/test_php_promoted_properties_visibility_modifiers
// origin: languages/php/tests/php/test_php_oop_constructor_promotion_readonly.rs
// vybe-test-mode: compile

class Service {
    public function __construct(
        private string $secretKey,
        protected string $endpoint,
        public int $timeout = 30
    ) {}
    
    public function getEndpoint(): string {
        return $this->endpoint;
    }
}

$s = new Service("key_123", "https://api.example.com");
echo $s->getEndpoint();
