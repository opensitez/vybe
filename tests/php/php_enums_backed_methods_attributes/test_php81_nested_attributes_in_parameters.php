<?php
// vybe-test: php/php_enums_backed_methods_attributes/test_php81_nested_attributes_in_parameters
// origin: languages/php/tests/php/test_php_enums_backed_methods_attributes.rs
// vybe-test-mode: compile

#[Attribute(Attribute::TARGET_PARAMETER)]
class Inject {
    public function __construct(public string $service) {}
}

class PaymentProcessor {
    public function __construct(
        #[Inject("db.connection")]
        public object $db
    ) {}
}

$rp = new ReflectionParameter([PaymentProcessor::class, "__construct"], "db");
echo count($rp->getAttributes(Inject::class));
