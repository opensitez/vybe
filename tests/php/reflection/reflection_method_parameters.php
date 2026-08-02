<?php
// vybe-test: php/reflection/reflection_method_parameters
// origin: languages/php/tests/php/test_reflection.rs
// vybe-test-mode: compile

function compute(int $x, float $y, string $label = 'default'): float {
    return $x + $y;
}
$rf = new ReflectionFunction('compute');
$params = $rf->getParameters();
echo count($params);
echo ':' . $params[2]->getDefaultValue();
