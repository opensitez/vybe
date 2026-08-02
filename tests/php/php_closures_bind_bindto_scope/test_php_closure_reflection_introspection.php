<?php
// vybe-test: php/php_closures_bind_bindto_scope/test_php_closure_reflection_introspection
// origin: languages/php/tests/php/test_php_closures_bind_bindto_scope.rs
// vybe-test-mode: compile

$fn = function(int $a, string $b = "default"): string {
    return $b . $a;
};

$rc = new ReflectionFunction($fn);
echo count($rc->getParameters());
