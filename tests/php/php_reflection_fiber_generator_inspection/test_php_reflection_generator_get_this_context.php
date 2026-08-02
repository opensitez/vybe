<?php
// vybe-test: php/php_reflection_fiber_generator_inspection/test_php_reflection_generator_get_this_context
// origin: languages/php/tests/php/test_php_reflection_fiber_generator_inspection.rs
// vybe-test-mode: compile

class GenRunner {
    public function run() {
        yield $this;
    }
}

$runner = new GenRunner();
$gen = $runner->run();
$gen->current();

$rg = new ReflectionGenerator($gen);
echo $rg->getThis() === $runner ? "THIS_BOUND_OK" : "FAIL";
