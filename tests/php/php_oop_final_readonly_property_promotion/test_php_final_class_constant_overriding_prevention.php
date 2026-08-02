<?php
// vybe-test: php/php_oop_final_readonly_property_promotion/test_php_final_class_constant_overriding_prevention
// origin: languages/php/tests/php/test_php_oop_final_readonly_property_promotion.rs
// vybe-test-mode: compile

class BaseConstants {
    final public const VERSION = "1.0.0";
}

class AppConstants extends BaseConstants {}

echo AppConstants::VERSION;
