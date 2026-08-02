<?php
// vybe-test: php/php_reflection_fiber_generator_inspection/test_php81_reflection_fiber_get_executing_file_and_line
// origin: languages/php/tests/php/test_php_reflection_fiber_generator_inspection.rs
// vybe-test-mode: compile

if (class_exists('Fiber') && class_exists('ReflectionFiber')) {
    $f = new Fiber(fn() => Fiber::suspend());
    $f->start();
    $rf = new ReflectionFiber($f);
    echo "File=" . strlen($rf->getExecutingFile()) . " Line=" . $rf->getExecutingLine();
}
