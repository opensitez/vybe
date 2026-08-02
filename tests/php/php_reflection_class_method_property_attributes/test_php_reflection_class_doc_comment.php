<?php
// vybe-test: php/php_reflection_class_method_property_attributes/test_php_reflection_class_doc_comment
// origin: languages/php/tests/php/test_php_reflection_class_method_property_attributes.rs
// vybe-test-mode: compile

/**
 * @Entity(table="products")
 */
class Product {}

$rc = new ReflectionClass(Product::class);
echo str_contains($rc->getDocComment(), "@Entity") ? "DOC_FOUND" : "NO_DOC";
