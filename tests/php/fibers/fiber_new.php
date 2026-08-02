<?php
// vybe-test: php/fibers/fiber_new
// origin: languages/php/tests/php/test_fibers.rs
// vybe-test-mode: compile

$fiber = new Fiber(function() {
    echo "Hello from fiber";
});
