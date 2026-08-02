<?php
// vybe-test: php/static_closures/static_closure_declared_with_keyword
// origin: languages/php/tests/php/test_static_closures.rs
// vybe-test-mode: compile

$fn = static function() {
    return 42;
};
echo $fn();
