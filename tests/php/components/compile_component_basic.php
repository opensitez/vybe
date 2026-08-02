<?php
// vybe-test: php/components/compile_component_basic
// origin: languages/php/tests/php/test_components.rs
// vybe-test-mode: compile

function greet($name) {
    return 'Hello ' . $name;
}
echo greet('World');
