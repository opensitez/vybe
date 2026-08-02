<?php
// vybe-test: php/php_oop_constructor_promotion_readonly/test_php_readonly_property_unassigned_error
// origin: languages/php/tests/php/test_php_oop_constructor_promotion_readonly.rs
// vybe-test-mode: compile

class Document {
    public readonly string $title;
}

$doc = new Document();
