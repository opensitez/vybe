<?php
// vybe-test: php/php_anonymous_classes_runtime_usage/test_php_anonymous_class_in_return_type_hint
// origin: languages/php/tests/php/test_php_anonymous_classes_runtime_usage.rs
// vybe-test-mode: compile

function createAnonymous(): object {
    return new class {
        public string $status = "active";
    };
}

$obj = createAnonymous();
echo $obj->status;
