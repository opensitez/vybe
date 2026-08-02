<?php
// vybe-test: php/php_enums_backed_methods_attributes/test_php81_repeatable_attribute
// origin: languages/php/tests/php/test_php_enums_backed_methods_attributes.rs
// vybe-test-mode: compile

#[Attribute(Attribute::TARGET_METHOD | Attribute::IS_REPEATABLE)]
class Middleware {
    public function __construct(public string $name) {}
}

class DashboardController {
    #[Middleware("auth")]
    #[Middleware("log")]
    public function index() {}
}

$rm = new ReflectionMethod(DashboardController::class, "index");
$attrs = $rm->getAttributes(Middleware::class);
echo count($attrs);
