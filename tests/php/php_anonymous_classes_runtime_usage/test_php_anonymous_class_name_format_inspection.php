<?php
// vybe-test: php/php_anonymous_classes_runtime_usage/test_php_anonymous_class_name_format_inspection
// origin: languages/php/tests/php/test_php_anonymous_classes_runtime_usage.rs
// vybe-test-mode: compile

$anon = new class {};
$className = get_class($anon);
echo str_contains($className, "class@anonymous") ? "ANONYMOUS_NAME_OK" : "NAMED";
