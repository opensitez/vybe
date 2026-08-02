<?php
// vybe-test: php/php_reflection_fiber_generator_inspection/test_php81_reflection_fiber_get_trace_callstack
// origin: languages/php/tests/php/test_php_reflection_fiber_generator_inspection.rs
// vybe-test-mode: compile

if (class_exists('Fiber') && class_exists('ReflectionFiber')) {
    $f = new Fiber(function() {
        Fiber::suspend();
    });
    $f->start();
    $rf = new ReflectionFiber($f);
    $trace = $rf->getTrace();
    echo is_array($trace) ? "FIBER_TRACE_OK" : "FAIL";
}
