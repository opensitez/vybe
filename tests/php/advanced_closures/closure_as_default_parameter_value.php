<?php
// vybe-test: php/advanced_closures/closure_as_default_parameter_value
// origin: languages/php/tests/php/test_advanced_closures.rs
// vybe-test-mode: compile

function transform(array $data, callable $fn = null): array {
    $fn ??= fn($x) => $x;
    return array_map($fn, $data);
}
$result = transform([1, 2, 3]);
echo count($result);
