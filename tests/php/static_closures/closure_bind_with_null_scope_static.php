<?php
// vybe-test: php/static_closures/closure_bind_with_null_scope_static
// origin: languages/php/tests/php/test_static_closures.rs
// vybe-test-mode: compile

$fn = static function() { return "static"; };
$bound = Closure::bind($fn, null, null);
echo $bound();
