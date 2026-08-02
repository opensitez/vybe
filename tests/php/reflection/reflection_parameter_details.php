<?php
// vybe-test: php/reflection/reflection_parameter_details
// origin: languages/php/tests/php/test_reflection.rs
// vybe-test-mode: compile

function create(string $name, int $age = 0, bool $active = true): void {}
$rf = new ReflectionFunction('create');
foreach ($rf->getParameters() as $p) {
    $opt = $p->isOptional() ? '?' : '!';
    echo $p->getName() . $opt . ' ';
}
