<?php
// vybe-test: php/php_anonymous_classes_runtime_usage/test_php_anonymous_class_with_attribute_annotation
// origin: languages/php/tests/php/test_php_anonymous_classes_runtime_usage.rs
// vybe-test-mode: compile

#[Attribute]
class ServiceMeta { public function __construct(public string $type) {} }

$service = new #[ServiceMeta("transient")] class {
    public function run() { return "ok"; }
};

$rc = new ReflectionClass($service);
echo count($rc->getAttributes(ServiceMeta::class));
