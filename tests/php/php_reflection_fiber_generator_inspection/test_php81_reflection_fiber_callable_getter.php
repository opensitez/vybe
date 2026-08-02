<?php
// vybe-test: php/php_reflection_fiber_generator_inspection/test_php81_reflection_fiber_callable_getter
// origin: languages/php/tests/php/test_php_reflection_fiber_generator_inspection.rs
// vybe-test-mode: compile

if (class_exists('Fiber') && class_exists('ReflectionFiber')) {
    $callable = function() { Fiber::suspend(); };
    $f = new Fiber($callable);
    $f->start();
    $rf = new ReflectionFiber($f);
    echo is_callable($rf->getCallable()) ? "CALLABLE_OK" : "FAIL";
}
